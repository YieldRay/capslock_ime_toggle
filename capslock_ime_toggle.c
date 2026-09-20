#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif
#include <windows.h>
#include <stdio.h>
#include <shellscalingapi.h> // 用于 SetProcessDPIAware
#include <shellapi.h>
#include <imm.h>
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "imm32.lib")

#define MUTEX_NAME L"Global\\CapsLockImeToggleMutex" // 全局互斥体名，防止多开
#define WM_TRAYICON (WM_USER + 1)                    // 托盘消息ID
#define WM_TOGGLE_IME (WM_USER + 2)                  // 异步切换输入法消息
#define TASK_NAME L"CapsLockImeToggle"              // 计划任务名称
#define SCHTASKS_EXECUTION_FAILED ((DWORD)-1)

#ifdef UNICODE
#define _tWinMain wWinMain
#else
#define _tWinMain WinMain
#endif

volatile BOOL g_allowNextCaps = 0; // 标志：是否允许下一个CapsLock事件通过钩子（用于托盘菜单手动切换大小写）
volatile BOOL g_capsLockDown = FALSE; // 标志：CapsLock按键是否已经处理，防止长按重复切换
HHOOK g_hHook = NULL;              // 全局低级键盘钩子句柄，用于拦截CapsLock按键
NOTIFYICONDATAW nid;               // 系统托盘图标数据结构
HANDLE g_hMutex = NULL;            // 全局互斥体句柄，防止程序多开
HWND g_hMainWnd = NULL;             // 隐藏窗口句柄，用于异步切换输入法

typedef enum
{
    AUTO_START_DISABLED,
    AUTO_START_ENABLED,
    AUTO_START_UNKNOWN
} AutoStartState;

// 判断当前进程是否为管理员
BOOL IsRunAsAdmin()
{
    BOOL isAdmin = FALSE;
    HANDLE hToken = NULL;
    if (OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
    {
        TOKEN_ELEVATION elevation;
        DWORD dwSize = sizeof(TOKEN_ELEVATION);
        if (GetTokenInformation(hToken, TokenElevation, &elevation, sizeof(elevation), &dwSize))
        {
            isAdmin = elevation.TokenIsElevated;
        }
        CloseHandle(hToken);
    }
    return isAdmin;
}

// 静默执行 schtasks.exe，返回进程退出码
DWORD RunSchtasks(LPCWSTR args)
{
    WCHAR schtasksPath[MAX_PATH];
    WCHAR cmdLine[2048];
    UINT systemDirectoryLength = GetSystemDirectoryW(schtasksPath, MAX_PATH);
    if (!systemDirectoryLength || systemDirectoryLength + lstrlenW(L"\\schtasks.exe") >= MAX_PATH)
    {
        return SCHTASKS_EXECUTION_FAILED;
    }
    lstrcatW(schtasksPath, L"\\schtasks.exe");
    swprintf(cmdLine, sizeof(cmdLine) / sizeof(WCHAR), L"\"%s\" %s", schtasksPath, args);

    STARTUPINFOW si = {0};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi = {0};
    DWORD exitCode = SCHTASKS_EXECUTION_FAILED;
    if (CreateProcessW(schtasksPath, cmdLine, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi))
    {
        DWORD waitResult = WaitForSingleObject(pi.hProcess, 10000);
        if (waitResult == WAIT_OBJECT_0)
            GetExitCodeProcess(pi.hProcess, &exitCode);
        else if (waitResult == WAIT_TIMEOUT)
            TerminateProcess(pi.hProcess, ERROR_TIMEOUT);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
    return exitCode;
}

// 通过 runas 提权执行 schtasks，仅在创建或删除任务时需要授权
DWORD RunSchtasksElevated(LPCWSTR args)
{
    WCHAR schtasksPath[MAX_PATH];
    UINT systemDirectoryLength = GetSystemDirectoryW(schtasksPath, MAX_PATH);
    if (!systemDirectoryLength || systemDirectoryLength + lstrlenW(L"\\schtasks.exe") >= MAX_PATH)
    {
        return SCHTASKS_EXECUTION_FAILED;
    }
    lstrcatW(schtasksPath, L"\\schtasks.exe");

    SHELLEXECUTEINFOW sei = {0};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb = L"runas";
    sei.lpFile = schtasksPath;
    sei.lpParameters = args;
    sei.nShow = SW_HIDE;
    DWORD exitCode = SCHTASKS_EXECUTION_FAILED;
    if (ShellExecuteExW(&sei) && sei.hProcess)
    {
        DWORD waitResult = WaitForSingleObject(sei.hProcess, 30000);
        if (waitResult == WAIT_OBJECT_0)
            GetExitCodeProcess(sei.hProcess, &exitCode);
        else if (waitResult == WAIT_TIMEOUT)
            TerminateProcess(sei.hProcess, ERROR_TIMEOUT);
        CloseHandle(sei.hProcess);
    }
    return exitCode;
}

DWORD RunSchtasksAuto(LPCWSTR args)
{
    if (IsRunAsAdmin())
        return RunSchtasks(args);
    return RunSchtasksElevated(args);
}

AutoStartState GetAutoStartState()
{
    DWORD result = RunSchtasks(L"/query /tn \"" TASK_NAME L"\"");
    if (result == 0)
        return AUTO_START_ENABLED;
    if (result == SCHTASKS_EXECUTION_FAILED)
        return AUTO_START_UNKNOWN;
    return AUTO_START_DISABLED;
}

DWORD EnableAutoStart()
{
    WCHAR exePath[MAX_PATH];
    DWORD exePathLength = GetModuleFileNameW(NULL, exePath, MAX_PATH);
    if (!exePathLength || exePathLength >= MAX_PATH)
        return SCHTASKS_EXECUTION_FAILED;

    WCHAR args[MAX_PATH * 2 + 128];
    swprintf(args, sizeof(args) / sizeof(WCHAR),
             L"/create /tn \"%s\" /tr \"\\\"%s\\\"\" /sc onlogon /it /rl highest /f",
             TASK_NAME, exePath);
    return RunSchtasksAuto(args);
}

DWORD DisableAutoStart()
{
    return RunSchtasksAuto(L"/delete /tn \"" TASK_NAME L"\" /f");
}

void ShowAutoStartFailureMessage(HWND hwnd, AutoStartState state, DWORD result)
{
    WCHAR message[256];
    if (result == SCHTASKS_EXECUTION_FAILED)
    {
        lstrcpyW(message, L"无法启动任务计划程序或用户取消了UAC授权。\n请重试并确认允许管理员授权。");
    }
    else
    {
        swprintf(message, sizeof(message) / sizeof(WCHAR), L"任务计划程序操作失败（返回码：%lu）。", result);
    }

    if (state == AUTO_START_ENABLED)
    {
        swprintf(message + lstrlenW(message), sizeof(message) / sizeof(WCHAR) - lstrlenW(message), L"\n当前仍已设置开机自启。");
    }
    else if (state == AUTO_START_DISABLED)
    {
        swprintf(message + lstrlenW(message), sizeof(message) / sizeof(WCHAR) - lstrlenW(message), L"\n当前仍未设置开机自启。");
    }
    else
    {
        swprintf(message + lstrlenW(message), sizeof(message) / sizeof(WCHAR) - lstrlenW(message), L"\n无法确认当前开机自启状态。");
    }
    MessageBoxW(hwnd, message, L"错误", MB_OK | MB_ICONERROR);
}

void ToggleAutoStart(HWND hwnd)
{
    AutoStartState initialState = GetAutoStartState();
    if (initialState == AUTO_START_UNKNOWN)
    {
        MessageBoxW(hwnd, L"无法查询当前开机自启状态，操作未执行。", L"错误", MB_OK | MB_ICONERROR);
        return;
    }

    BOOL enable = initialState == AUTO_START_DISABLED;
    DWORD operationResult = enable ? EnableAutoStart() : DisableAutoStart();
    if (operationResult == 0)
    {
        if (enable)
        {
            MessageBoxW(hwnd, L"已设置开机自启！\n登录时将自动以管理员权限运行，不会再弹出UAC窗口。", L"提示", MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBoxW(hwnd, L"已取消开机自启！", L"提示", MB_OK | MB_ICONINFORMATION);
        }
        return;
    }

    AutoStartState finalState = GetAutoStartState();
    ShowAutoStartFailureMessage(hwnd, finalState, operationResult);
}

// 添加系统托盘图标
void AddTrayIcon(HWND hwnd)
{
    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIconW(GetModuleHandleW(NULL), MAKEINTRESOURCEW(101)); // 101是IDI_TRAYICON的资源ID值，定义在icon.rc中
    if (IsRunAsAdmin())
    {
        lstrcpyW(nid.szTip, L"CapsLock IME Toggle (管理员)");
    }
    else
    {
        lstrcpyW(nid.szTip, L"CapsLock IME Toggle");
    }
    if (!Shell_NotifyIconW(NIM_ADD, &nid))
    {
        MessageBoxW(hwnd, L"添加托盘图标失败！", L"错误", MB_OK | MB_ICONERROR);
    }
}

// 移除系统托盘图标
void RemoveTrayIcon()
{
    Shell_NotifyIconW(NIM_DELETE, &nid);
}

// 显示托盘右键菜单
void ShowTrayMenu(HWND hwnd)
{
    POINT pt;
    GetCursorPos(&pt);
    HMENU hMenu = CreatePopupMenu();
    AutoStartState autoStartState = GetAutoStartState();
    AppendMenuW(hMenu, MF_STRING, 100, L"切换大小写");
    AppendMenuW(hMenu, MF_STRING | (autoStartState == AUTO_START_ENABLED ? MF_CHECKED : 0), 200, L"切换开机自启（管理员）");
    AppendMenuW(hMenu, MF_STRING, 300, L"输入法设置");
    AppendMenuW(hMenu, MF_STRING, 400, L"以管理员模式重启");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, 1, L"退出");
    SetForegroundWindow(hwnd);
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
    switch (cmd)
    {
    case 1: // 退出程序
        PostMessage(hwnd, WM_CLOSE, 0, 0);
        break;
    case 100: // 切换大小写，允许下一个CapsLock事件通过钩子
        g_allowNextCaps = TRUE;
        keybd_event(VK_CAPITAL, 0, 0, 0);
        keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0);
        break;
    case 200: // 切换开机自启（任务计划程序方式）
        ToggleAutoStart(hwnd);
        break;
    case 300: // 打开Windows 10/11输入法设置
    {
        HINSTANCE hRet = ShellExecuteW(NULL, L"open", L"ms-settings:regionlanguage", NULL, NULL, SW_SHOWNORMAL);
        if ((INT_PTR)hRet <= 32)
        {
            MessageBoxW(hwnd, L"无法打开输入法设置！", L"错误", MB_OK | MB_ICONERROR);
        }
        break;
    }
    case 400: // 以管理员模式重启当前进程
    {
        if (IsRunAsAdmin())
        {
            {
                int ret = MessageBoxW(hwnd, L"当前已是管理员，无需重启！\n是否强制以管理员模式重启？", L"提示", MB_OKCANCEL | MB_ICONQUESTION);
                if (ret == IDOK)
                {
                    WCHAR exePath[MAX_PATH];
                    if (GetModuleFileNameW(NULL, exePath, MAX_PATH))
                    {
                        if (g_hMutex)
                        {
                            CloseHandle(g_hMutex); // 先释放互斥体
                            g_hMutex = NULL;
                        }
                        HINSTANCE restartResult = ShellExecuteW(NULL, L"runas", exePath, NULL, NULL, SW_SHOWNORMAL);
                        if ((INT_PTR)restartResult > 32)
                            PostMessage(hwnd, WM_CLOSE, 0, 0); // 关闭当前进程
                        else
                            MessageBoxW(hwnd, L"管理员模式重启失败，当前程序仍在运行。", L"错误", MB_OK | MB_ICONERROR);
                    }
                    else
                    {
                        MessageBoxW(hwnd, L"获取程序路径失败，无法重启！", L"错误", MB_OK | MB_ICONERROR);
                    }
                }
                // 取消则什么都不做
                break;
            }
        }
        WCHAR exePath[MAX_PATH];
        if (GetModuleFileNameW(NULL, exePath, MAX_PATH))
        {
            if (g_hMutex)
            {
                CloseHandle(g_hMutex); // 先释放互斥体
                g_hMutex = NULL;
            }
            HINSTANCE restartResult = ShellExecuteW(NULL, L"runas", exePath, NULL, NULL, SW_SHOWNORMAL);
            if ((INT_PTR)restartResult > 32)
                PostMessage(hwnd, WM_CLOSE, 0, 0); // 关闭当前进程
            else
                MessageBoxW(hwnd, L"管理员模式重启失败，当前程序仍在运行。", L"错误", MB_OK | MB_ICONERROR);
        }
        else
        {
            MessageBoxW(hwnd, L"获取程序路径失败，无法重启！", L"错误", MB_OK | MB_ICONERROR);
        }
        break;
    }
    default:
        break;
    }
}

// 切换当前激活窗口的输入法中英文状态
// 兼容微软拼音和主流第三方输入法
void ToggleImeConversion(HWND hwnd)
{
    DWORD threadId = GetWindowThreadProcessId(hwnd, NULL);
    HKL hKL = GetKeyboardLayout(threadId);
    // 微软拼音输入法的HKL前缀通常为0xE00xxxxxx
    // 0x0804E00x: 微软拼音，0x0804xxxx: 其他中文输入法
    // 根据目标窗口线程的键盘布局判断微软拼音或兼容输入法
    if ((UINT_PTR)hKL == 0xE00E0804 || ((UINT_PTR)hKL & 0xF0000000) == 0xE0000000)
    {
        HIMC hIMC = ImmGetContext(hwnd);
        if (hIMC)
        {
            DWORD conv, sent;
            if (ImmGetConversionStatus(hIMC, &conv, &sent))
            {
                // 切换中英文（IME_CMODE_NATIVE）
                conv ^= IME_CMODE_NATIVE;
                ImmSetConversionStatus(hIMC, conv, sent);
            }
            ImmReleaseContext(hwnd, hIMC);
        }
        else
        {
            MessageBoxW(hwnd, L"获取输入法上下文失败！", L"错误", MB_OK | MB_ICONERROR);
        }
    }
    else
    {
        // 其他输入法（如搜狗、QQ等），模拟Ctrl+Shift+Space快捷键
        // 先激活目标窗口
        if (SetForegroundWindow(hwnd))
        {
            keybd_event(VK_CONTROL, 0, 0, 0);
            keybd_event(VK_SHIFT, 0, 0, 0);
            keybd_event(VK_SPACE, 0, 0, 0);
            keybd_event(VK_SPACE, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
        }
        else
        {
            MessageBoxW(hwnd, L"无法激活目标窗口，输入法切换失败！", L"错误", MB_OK | MB_ICONERROR);
        }
    }
}

// 全局低级键盘钩子回调，拦截CapsLock
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION)
    {
        KBDLLHOOKSTRUCT *p = (KBDLLHOOKSTRUCT *)lParam;
        if ((wParam == WM_KEYUP || wParam == WM_SYSKEYUP) && p->vkCode == VK_CAPITAL)
        {
            g_capsLockDown = FALSE;
            return CallNextHookEx(g_hHook, nCode, wParam, lParam);
        }
        if (wParam == WM_KEYDOWN && p->vkCode == VK_CAPITAL)
        {
            if (g_capsLockDown)
                return 1;

            // 菜单触发时允许本次CapsLock事件通过
            if (g_allowNextCaps)
            {
                g_allowNextCaps = FALSE;
                g_capsLockDown = TRUE;
                return CallNextHookEx(g_hHook, nCode, wParam, lParam);
            }
            // 检查是否有修饰键（Ctrl/Shift/Alt/Win）按下，只有无修饰键时才拦截
            SHORT ctrl = GetAsyncKeyState(VK_CONTROL);
            SHORT shift = GetAsyncKeyState(VK_SHIFT);
            SHORT alt = GetAsyncKeyState(VK_MENU);
            SHORT win = GetAsyncKeyState(VK_LWIN) | GetAsyncKeyState(VK_RWIN);
            if (!(ctrl & 0x8000) && !(shift & 0x8000) && !(alt & 0x8000) && !(win & 0x8000))
            {
                HWND targetWindow = GetForegroundWindow();
                if (targetWindow && g_hMainWnd &&
                    PostMessageW(g_hMainWnd, WM_TOGGLE_IME, (WPARAM)targetWindow, 0))
                {
                    g_capsLockDown = TRUE;
                    return 1;
                }
                // 不传递CapsLock按键，防止系统CapsLock状态改变
                return CallNextHookEx(g_hHook, nCode, wParam, lParam);
            }
            // 有修饰键时不拦截，允许系统处理
        }
    }
    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

// 主窗口过程，处理托盘消息和退出
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP)
        {
            ShowTrayMenu(hwnd);
        }
        else if (lParam == WM_LBUTTONDBLCLK)
        {
            MessageBoxW(hwnd, L"CapsLock IME Toggle 正在运行\n右键托盘图标可退出", L"提示", MB_OK | MB_ICONINFORMATION);
        }
        break;
    case WM_TOGGLE_IME:
        if (IsWindow((HWND)wParam))
            ToggleImeConversion((HWND)wParam);
        break;
    case WM_DESTROY:
        RemoveTrayIcon();
        if (g_hHook)
            UnhookWindowsHookEx(g_hHook);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// 程序主入口，初始化互斥体、窗口、钩子和消息循环
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    SetProcessDPIAware();
    // 创建互斥体，防止多开
    g_hMutex = CreateMutexW(NULL, FALSE, MUTEX_NAME);
    if (g_hMutex == NULL)
    {
        MessageBoxW(NULL, L"创建程序互斥体失败，程序无法启动。", L"错误", MB_OK | MB_ICONERROR);
        return 1;
    }
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        MessageBoxW(NULL, L"CapsLock IME Toggle 已经在运行！", L"提示", MB_OK | MB_ICONINFORMATION);
        if (g_hMutex)
            CloseHandle(g_hMutex);
        return 0;
    }
    // 注册窗口类
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CapsLockImeToggleClass";
    wc.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(101));
    wc.hCursor = LoadCursorW(NULL, MAKEINTRESOURCEW(IDC_ARROW));
    if (!RegisterClassW(&wc))
    {
        MessageBoxW(NULL, L"窗口类注册失败！", L"错误", MB_OK | MB_ICONERROR);
        if (g_hMutex)
            CloseHandle(g_hMutex);
        return 1;
    }

    // 创建隐藏窗口（不显示在任务栏，仅用于托盘消息）
    HWND hwnd = CreateWindowW(
        L"CapsLockImeToggleClass",
        L"CapsLock IME Toggle",
        WS_OVERLAPPED | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 320, 100,
        NULL, NULL, hInstance, NULL);
    if (!hwnd)
    {
        MessageBoxW(NULL, L"窗口创建失败！", L"错误", MB_OK | MB_ICONERROR);
        if (g_hMutex)
            CloseHandle(g_hMutex);
        return 1;
    }
    g_hMainWnd = hwnd;

    // 添加托盘图标
    AddTrayIcon(hwnd);

    // 设置全局低级键盘钩子
    g_hHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);
    if (!g_hHook)
    {
        MessageBoxW(hwnd, L"无法安装键盘钩子", L"错误", MB_ICONERROR);
        RemoveTrayIcon();
        DestroyWindow(hwnd);
        if (g_hMutex)
            CloseHandle(g_hMutex);
        return 1;
    }

    // 消息循环
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    if (g_hMutex)
        CloseHandle(g_hMutex);
    return 0;
}