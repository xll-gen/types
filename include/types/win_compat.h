#pragma once

#include <algorithm>
#include <cstdarg>
#include <cstdint>
#include <cstring>
#include <cwchar>
#include <string>

#ifdef _WIN32
#error "win_compat.h should not be included on Windows"
#endif

// Basic Types
typedef int32_t INT32;
typedef wchar_t WCHAR;
typedef wchar_t XCHAR;
typedef int32_t RW;
typedef int32_t COL;
typedef uintptr_t DWORD_PTR;
typedef uint16_t WORD;
typedef uint8_t BYTE;
typedef char* LPSTR;
typedef void* HANDLE;
typedef void VOID;
typedef void* HWND;
typedef void* HINSTANCE;
typedef void* HMODULE;
typedef struct tagPOINT {
    int32_t x;
    int32_t y;
} POINT;

typedef uint32_t DWORD;

#define __stdcall
#define pascal
#define PASCAL
#define _cdecl
#define CALLBACK

#define __forceinline inline
#define __declspec(x)

#define MAX_PATH 260

// Mock Windows API constants/functions
#define CP_UTF8 65001

int MultiByteToWideChar(unsigned int CodePage, unsigned long dwFlags, const char* lpMultiByteStr, int cbMultiByte,
                        wchar_t* lpWideCharStr, int cchWideChar);
int WideCharToMultiByte(unsigned int CodePage, unsigned long dwFlags, const wchar_t* lpWideCharStr, int cchWideChar,
                        char* lpMultiByteStr, int cbMultiByte, const char* lpDefaultChar, int* lpUsedDefaultChar);

unsigned long GetModuleFileNameW(void* hModule, wchar_t* lpFilename, unsigned long nSize);
void* GetModuleHandle(const char* lpModuleName);
void* GetProcAddress(void* hModule, const char* lpProcName);

extern void* g_hModule;
