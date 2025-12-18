#ifndef _WIN32
#include "types/win_compat.h"

#include <codecvt>
#include <cstring>
#include <locale>
#include <vector>

// Global module handle definition
void* g_hModule = nullptr;

int MultiByteToWideChar(unsigned int CodePage, unsigned long dwFlags, const char* lpMultiByteStr, int cbMultiByte,
                        wchar_t* lpWideCharStr, int cchWideChar) {
    std::string str;
    if (cbMultiByte == -1)
        str = lpMultiByteStr;
    else
        str.assign(lpMultiByteStr, cbMultiByte);

    try {
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::wstring wstr = converter.from_bytes(str);

        int resLen = (int)wstr.length();
        if (cbMultiByte == -1) resLen += 1;  // Null terminator

        if (cchWideChar == 0) {
            return resLen;
        }

        int count = std::min(resLen, cchWideChar);
        if (cbMultiByte == -1 && count > 0 && count == resLen) {
            // Include null
            std::memcpy(lpWideCharStr, wstr.data(), (count - 1) * sizeof(wchar_t));
            lpWideCharStr[count - 1] = 0;
        } else if (count > 0) {
            std::memcpy(lpWideCharStr, wstr.data(), count * sizeof(wchar_t));
        }
        return count;
    } catch (...) {
        return 0;
    }
}

int WideCharToMultiByte(unsigned int CodePage, unsigned long dwFlags, const wchar_t* lpWideCharStr, int cchWideChar,
                        char* lpMultiByteStr, int cbMultiByte, const char* lpDefaultChar, int* lpUsedDefaultChar) {
    std::wstring wstr;
    if (cchWideChar == -1)
        wstr = lpWideCharStr;
    else
        wstr.assign(lpWideCharStr, cchWideChar);

    try {
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::string str = converter.to_bytes(wstr);

        int resLen = (int)str.length();
        if (cchWideChar == -1) resLen += 1;

        if (cbMultiByte == 0) {
            return resLen;
        }

        int count = std::min(resLen, cbMultiByte);
        if (cchWideChar == -1 && count > 0 && count == resLen) {
            std::memcpy(lpMultiByteStr, str.data(), count - 1);
            lpMultiByteStr[count - 1] = 0;
        } else if (count > 0) {
            std::memcpy(lpMultiByteStr, str.data(), count);
        }
        return count;
    } catch (...) {
        return 0;
    }
}

unsigned long GetModuleFileNameW(void* hModule, wchar_t* lpFilename, unsigned long nSize) {
    if (nSize > 0 && lpFilename) lpFilename[0] = 0;
    return 0;
}

void* GetModuleHandle(const char* lpModuleName) { return nullptr; }
void* GetProcAddress(void* hModule, const char* lpProcName) { return nullptr; }

#endif
