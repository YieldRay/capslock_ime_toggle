CC=gcc
WINDRES=windres
CFLAGS=-mwindows -municode
LDFLAGS=-limm32
TARGET=capslock_ime_toggle.exe
SRC=capslock_ime_toggle.c
RES=icon.res

all: $(TARGET)

$(RES): icon.rc icon.ico
	$(WINDRES) --input=icon.rc --output=$@ --output-format=coff --target=pe-x86-64

$(TARGET): $(SRC) $(RES)
	$(CC) $(CFLAGS) -o $@ $(SRC) $(RES) $(LDFLAGS)

clean:
	del /Q $(TARGET) $(RES)
