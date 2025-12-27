#pragma once
#ifdef _WIN32
#include <windows.h>
#else
#include "types/win_compat.h"
#endif
#include <string>
#include <vector>
#include "xlcall.h"

typedef wchar_t XLL_PASCAL_STRING;

// Global Module Handle (extern)
extern HINSTANCE g_hModule;

// Registration Helpers
LPXLOPER12 TempStr12(const wchar_t* txt);
LPXLOPER12 TempInt12(int val);

// String Conversion Helpers
std::wstring StringToWString(const std::string& str);
std::string WideToUtf8(const std::wstring& wstr);
std::wstring ConvertToWString(const char* str);
std::string ConvertExcelString(const wchar_t* wstr);
void Utf8ToExcelString(const char* utf8, XCHAR*& outStr);

// Cell Helper
bool IsSingleCell(LPXLOPER12 pxRef);

// Path Helper
std::wstring GetXllDir();

// Debug Logging
void SetDebugFlag(bool enabled);
bool GetDebugFlag();
void DebugLog(const char* fmt, ...);

// Include ScopedXLOPER12 helper
#include "types/ScopedXLOPER12.h"
