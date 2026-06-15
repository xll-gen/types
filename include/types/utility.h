#pragma once
#include <windows.h>
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
std::wstring PascalToWString(const wchar_t* pstr);
std::string ConvertExcelString(const wchar_t* wstr);
void Utf8ToExcelString(const char* utf8, XCHAR*& outStr);

// Cell Helper
bool IsSingleCell(LPXLOPER12 pxRef);

// True if a number-format code displays its value as a date/time, i.e. it
// contains an unescaped date/time token (y/m/d/h/s) outside quoted literals
// ("...") and bracketed sections ([Red], [$-409], [h]). "General" and pure
// numeric codes return false. Used to decide whether a date cell already has a
// suitable format (skip) or needs the default applied.
bool IsDateLikeFormat(const std::wstring& fmt);

// Path Helper
std::wstring GetXllDir();

// Debug Logging
void SetDebugFlag(bool enabled);
bool GetDebugFlag();
void DebugLog(const char* fmt, ...);

// Include ScopedXLOPER12 helper
#include "types/ScopedXLOPER12.h"
