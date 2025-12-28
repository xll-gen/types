#include "types/utility.h"
#include <vector>
#include <stdexcept>
#include <limits>
#include <cstdarg>
#include <cstdio>
#include <cstring>
#include "types/mem.h" // For TempStr12/TempInt12 if needed (actually they use thread local)

#ifdef __linux__
#include <unistd.h>
#include <limits.h>
#endif

// Limit strings to 10MB to prevent DoS
static const size_t MAX_STRING_SIZE = 10 * 1024 * 1024;

std::wstring StringToWString(const std::string& str) {
    if (str.empty()) return std::wstring();

    if (str.size() > MAX_STRING_SIZE) {
        throw std::length_error("String too long for conversion");
    }

    if (str.size() > (size_t)std::numeric_limits<int>::max()) {
        throw std::length_error("String too long for conversion");
    }

    int size_needed = MultiByteToWideChar(65001, 0, &str[0], (int)str.size(), NULL, 0);
    if (size_needed <= 0) return std::wstring();

    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(65001, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

std::string WideToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return "";

    // Explicit check as per Bug 15 mitigation
    if (wstr.size() > (size_t)std::numeric_limits<int>::max()) {
         throw std::length_error("Wide string too long (overflow)");
    }

    int size_needed = WideCharToMultiByte(65001, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    if (size_needed <= 0) return "";

    std::string strTo(size_needed, 0);
    WideCharToMultiByte(65001, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

std::wstring PascalToWString(const wchar_t* pstr) {
    if (!pstr) return L"";
    return std::wstring(pstr + 1, (size_t)pstr[0]);
}

std::wstring ConvertToWString(const char* str) {
    if (!str) return std::wstring();
    std::string s(str);
    return StringToWString(s);
}

std::string ConvertExcelString(const wchar_t* wstr) {
    if (!wstr) return "";

    // Excel strings are Pascal-like: First char is length
    size_t len = (size_t)wstr[0];

    if (len == 0) return "";

    // Skip the length prefix
    const wchar_t* actualStr = wstr + 1;

    // Safety check for huge strings (unlikely in Excel but possible in memory)
    if (len > MAX_STRING_SIZE) {
         throw std::length_error("String too long for conversion");
    }

    if (len > (size_t)std::numeric_limits<int>::max()) {
        throw std::length_error("String too long for conversion");
    }

    int size_needed = WideCharToMultiByte(65001, 0, actualStr, (int)len, NULL, 0, NULL, NULL);
    if (size_needed <= 0) return "";

    std::string strTo(size_needed, 0);
    WideCharToMultiByte(65001, 0, actualStr, (int)len, &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

void Utf8ToExcelString(const char* utf8, XCHAR*& outStr) {
    if (!utf8) {
        outStr = new XCHAR[2];
        outStr[0] = 0;
        outStr[1] = 0;
        return;
    }

    size_t realLen = strlen(utf8);

    // Optimization/Security: Excel 12 strings are limited to 32767 chars.
    // If the input is huge, we don't need to convert all of it just to truncate it.
    // 32767 chars can be at most ~132KB (4 bytes per char) in UTF-8.
    // We clamp the input to a safe margin (200,000 bytes) to prevent allocating
    // huge temporary buffers (e.g. 20MB) for strings that will be truncated anyway.
    // This also prevents integer overflow when casting size_t to int.
    if (realLen > 200000) {
        realLen = 200000;
    }
    int utf8Len = static_cast<int>(realLen);

    // CP_UTF8 = 65001
    // Optimization: Try to convert using a stack buffer first to avoid double API call.
    // Most Excel strings are small.
    XCHAR stackBuf[256];
    int needed = 0;

    if (utf8Len < 256) {
        needed = MultiByteToWideChar(65001, 0, utf8, utf8Len, stackBuf, 256);
    }

    if (needed > 0) {
        // Successful conversion to stack buffer
        outStr = new XCHAR[needed + 2];
        std::memcpy(outStr + 1, stackBuf, needed * sizeof(XCHAR));
        outStr[0] = (XCHAR)needed;
        outStr[needed + 1] = 0;
    } else {
        // Fallback to double-call (or string too long for stack buffer)
        needed = MultiByteToWideChar(65001, 0, utf8, utf8Len, NULL, 0);

        // Safety check: Don't allocate huge memory for strings.
        // Use MAX_STRING_SIZE for consistency (Issue 19)
        if (needed < 0 || (size_t)needed > MAX_STRING_SIZE || (size_t)needed > SIZE_MAX / sizeof(XCHAR) - 2) {
            outStr = new XCHAR[2];
            outStr[0] = 0;
            outStr[1] = 0;
        } else {
            if (needed > 32767) {
                // Manual memory management for temp since we don't have ScopeGuard header here or want to keep it simple
                // We use std::vector to avoid manual delete[]
                std::vector<XCHAR> tempVec(needed + 2);

                MultiByteToWideChar(65001, 0, utf8, utf8Len, tempVec.data() + 1, needed);

                outStr = new XCHAR[32767 + 2];
                std::memcpy(outStr + 1, tempVec.data() + 1, 32767 * sizeof(XCHAR));

                outStr[0] = 32767;
                outStr[32767 + 1] = 0;
            } else {
                outStr = new XCHAR[needed + 2];
                if (needed > 0) {
                    MultiByteToWideChar(65001, 0, utf8, utf8Len, outStr + 1, needed);
                }
                outStr[0] = (XCHAR)needed;
                outStr[needed + 1] = 0;
            }
        }
    }
}

// Restoring TempStr12 and TempInt12 as they were removed
LPXLOPER12 TempStr12(const wchar_t* txt) {
    static thread_local XLOPER12 xOp[10];
    static thread_local int i = 0;
    i = (i + 1) % 10;
    LPXLOPER12 op = &xOp[i];

    op->xltype = xltypeStr;
    static thread_local wchar_t strBuf[10][256];
    size_t len = 0;
    if (txt) len = wcslen(txt);
    if (len > 255) len = 255;

    strBuf[i][0] = (wchar_t)len;
    if (len > 0) wmemcpy(strBuf[i]+1, txt, len);

    op->val.str = strBuf[i];
    return op;
}

LPXLOPER12 TempInt12(int val) {
    static thread_local XLOPER12 xOp[10];
    static thread_local int i = 0;
    i = (i + 1) % 10;
    LPXLOPER12 op = &xOp[i];
    op->xltype = xltypeInt;
    op->val.w = val;
    return op;
}

bool IsSingleCell(LPXLOPER12 pxRef) {
    if (!pxRef) return false;
    if (pxRef->xltype & xltypeSRef) {
        return (pxRef->val.sref.ref.rwFirst == pxRef->val.sref.ref.rwLast) &&
               (pxRef->val.sref.ref.colFirst == pxRef->val.sref.ref.colLast);
    }
    if (pxRef->xltype & xltypeRef) {
        return (pxRef->val.mref.lpmref->count == 1) &&
               (pxRef->val.mref.lpmref->reftbl[0].rwFirst == pxRef->val.mref.lpmref->reftbl[0].rwLast) &&
               (pxRef->val.mref.lpmref->reftbl[0].colFirst == pxRef->val.mref.lpmref->reftbl[0].colLast);
    }
    return false;
}

std::wstring GetXllDir() {
#ifdef __linux__
    char result[PATH_MAX];
    ssize_t count = readlink("/proc/self/exe", result, PATH_MAX);
    if (count != -1) {
        std::string p(result, count);
        size_t pos = p.find_last_of('/');
        if (pos != std::string::npos) {
            return StringToWString(p.substr(0, pos));
        }
    }
    return L".";
#else
    wchar_t path[MAX_PATH];
    if (GetModuleFileNameW(g_hModule, path, MAX_PATH) == 0) return L"";
    std::wstring p(path);
    size_t pos = p.find_last_of(L"\\/");
    if (pos != std::wstring::npos) {
        return p.substr(0, pos);
    }
    return L".";
#endif
}

// Debug Logging
static bool g_debug = false;

void SetDebugFlag(bool debug) {
    g_debug = debug;
}

bool GetDebugFlag() {
    return g_debug;
}

void DebugLog(const char* fmt, ...) {
    if (!g_debug) return;

    va_list args;
    va_start(args, fmt);
    char buffer[1024];
    vsnprintf(buffer, sizeof(buffer), fmt, args);
    va_end(args);

#ifdef _WIN32
    OutputDebugStringA(buffer);
    OutputDebugStringA("\n");
#else
    fprintf(stderr, "[DEBUG] %s\n", buffer);
#endif
}
