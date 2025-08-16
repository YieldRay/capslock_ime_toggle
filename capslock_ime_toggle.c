#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif
#include <windows.h>
#include <shellscalingapi.h> // for SetProcessDPIAware
#include <shellapi.h>
#include <imm.h>
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "imm32.lib")

#define MUTEX_NAME L"Global\\CapsLockImeToggleMutex" // 全局互斥体名，防止多开
#define WM_TRAYICON (WM_USER + 1)                    // 托盘消息ID

#ifdef UNICODE
#define _tWinMain wWinMain
#else
#define _tWinMain WinMain
#endif

volatile BOOL g_allowNextCaps = 0;  // 标志：是否允许下一个CapsLock事件通过钩子（用于托盘菜单手动切换大小写）
HHOOK g_hHook = NULL;               // 全局低级键盘钩子句柄，用于拦截CapsLock按键
NOTIFYICONDATAW nid;                // 系统托盘图标数据结构
HANDLE g_hMutex = NULL;             // 全局互斥体句柄，防止程序多开

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
    AppendMenuW(hMenu, MF_STRING, 100, L"切换大小写");
    AppendMenuW(hMenu, MF_STRING, 200, L"切换开机自启");
    AppendMenuW(hMenu, MF_STRING, 300, L"输入法设置");
    AppendMenuW(hMenu, MF_STRING, 400, L"以管理员模式重启");
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, 1, L"退出");
    SetForegroundWindow(hwnd);
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
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
    case 200: // 切换开机自启
    {
        WCHAR exePath[MAX_PATH];
        if (GetModuleFileNameW(NULL, exePath, MAX_PATH))
        {
            HKEY hKey;
            LONG res = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_SET_VALUE | KEY_QUERY_VALUE, &hKey);
            if (res == ERROR_SUCCESS)
            {
                // 检查是否已设置
                WCHAR val[MAX_PATH] = {0};
                DWORD type = 0, size = sizeof(val);
                res = RegQueryValueExW(hKey, L"CapsLockImeToggle", NULL, &type, (BYTE *)val, &size);
                if (res == ERROR_SUCCESS && type == REG_SZ && lstrcmpiW(val, exePath) == 0)
                {
                    // 已设置，执行取消
                    res = RegDeleteValueW(hKey, L"CapsLockImeToggle");
                    if (res == ERROR_SUCCESS)
                    {
                        MessageBoxW(hwnd, L"已取消开机自启！", L"提示", MB_OK | MB_ICONINFORMATION);
                    }
                    else
                    {
                        MessageBoxW(hwnd, L"取消开机自启失败，可能未设置或无权限！", L"错误", MB_OK | MB_ICONERROR);
                    }
                }
                else
                {
                    // 未设置，执行设置
                    res = RegSetValueExW(hKey, L"CapsLockImeToggle", 0, REG_SZ, (BYTE *)exePath, (lstrlenW(exePath) + 1) * sizeof(WCHAR));
                    if (res == ERROR_SUCCESS)
                    {
                        MessageBoxW(hwnd, L"已设置为开机自启！", L"提示", MB_OK | MB_ICONINFORMATION);
                    }
                    else
                    {
                        MessageBoxW(hwnd, L"设置开机自启失败！", L"错误", MB_OK | MB_ICONERROR);
                    }
                }
                RegCloseKey(hKey);
            }
            else
            {
                MessageBoxW(hwnd, L"无法访问注册表，操作失败！", L"错误", MB_OK | MB_ICONERROR);
            }
        }
        else
        {
            MessageBoxW(hwnd, L"获取程序路径失败！", L"错误", MB_OK | MB_ICONERROR);
        }
        break;
    }
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
            MessageBoxW(hwnd, L"当前已是管理员，无需重启！", L"提示", MB_OK | MB_ICONINFORMATION);
            break;
        }
        WCHAR exePath[MAX_PATH];
        if (GetModuleFileNameW(NULL, exePath, MAX_PATH))
        {
            if (g_hMutex)
            {
                CloseHandle(g_hMutex); // 先释放互斥体
                g_hMutex = NULL;
            }
            ShellExecuteW(NULL, L"runas", exePath, NULL, NULL, SW_SHOWNORMAL);
            PostMessage(hwnd, WM_CLOSE, 0, 0); // 关闭当前进程
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
    WCHAR imeName[32] = {0};
    GetKeyboardLayoutNameW(imeName);
    // 判断是否为微软拼音（或兼容IME）
    if ((UINT_PTR)hKL == 0xE00E0804 || wcsstr(imeName, L"E00") == imeName)
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
        if (wParam == WM_KEYDOWN && p->vkCode == VK_CAPITAL)
        {
            // 菜单触发时允许本次CapsLock事件通过
            if (g_allowNextCaps)
            {
                g_allowNextCaps = FALSE;
                return CallNextHookEx(g_hHook, nCode, wParam, lParam);
            }
            // 检查是否有修饰键（Ctrl/Shift/Alt/Win）按下，只有无修饰键时才拦截
            SHORT ctrl = GetAsyncKeyState(VK_CONTROL);
            SHORT shift = GetAsyncKeyState(VK_SHIFT);
            SHORT alt = GetAsyncKeyState(VK_MENU);
            SHORT win = GetAsyncKeyState(VK_LWIN) | GetAsyncKeyState(VK_RWIN);
            if (!(ctrl & 0x8000) && !(shift & 0x8000) && !(alt & 0x8000) && !(win & 0x8000))
            {
                HWND hwnd = GetForegroundWindow();
                if (hwnd)
                {
                    ToggleImeConversion(hwnd);
                }
                // 不传递CapsLock按键，防止系统CapsLock状态改变
                return 1;
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
    if (g_hMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS)
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