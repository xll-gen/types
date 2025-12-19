#ifndef _WIN32
#include "types/win_compat.h"
#include <locale>
#include <codecvt>
#include <vector>
#include <cstring>

// Global module handle definition
void* g_hModule = nullptr;

int MultiByteToWideChar([[maybe_unused]] unsigned int CodePage, [[maybe_unused]] unsigned long dwFlags, const char* lpMultiByteStr, int cbMultiByte, wchar_t* lpWideCharStr, int cchWideChar) {
    std::string str;
    if (cbMultiByte == -1) str = lpMultiByteStr;
    else str.assign(lpMultiByteStr, cbMultiByte);

    try {
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::wstring wstr = converter.from_bytes(str);

        int resLen = (int)wstr.length();
        if (cbMultiByte == -1) resLen += 1; // Null terminator

        if (cchWideChar == 0) {
            return resLen;
        }

        if (resLen > cchWideChar) {
            return 0;
        }

        if (cbMultiByte == -1) {
            // Include null
            std::memcpy(lpWideCharStr, wstr.data(), (resLen - 1) * sizeof(wchar_t));
            lpWideCharStr[resLen - 1] = 0;
        } else {
             std::memcpy(lpWideCharStr, wstr.data(), resLen * sizeof(wchar_t));
        }
        return resLen;
    } catch (...) {
        return 0;
    }
}

int WideCharToMultiByte([[maybe_unused]] unsigned int CodePage, [[maybe_unused]] unsigned long dwFlags, const wchar_t* lpWideCharStr, int cchWideChar, char* lpMultiByteStr, int cbMultiByte, [[maybe_unused]] const char* lpDefaultChar, [[maybe_unused]] int* lpUsedDefaultChar) {
    std::wstring wstr;
    if (cchWideChar == -1) wstr = lpWideCharStr;
    else wstr.assign(lpWideCharStr, cchWideChar);

    try {
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::string str = converter.to_bytes(wstr);

        int resLen = (int)str.length();
        if (cchWideChar == -1) resLen += 1;

        if (cbMultiByte == 0) {
            return resLen;
        }

        if (resLen > cbMultiByte) {
            return 0;
        }

        if (cchWideChar == -1) {
            std::memcpy(lpMultiByteStr, str.data(), resLen - 1);
            lpMultiByteStr[resLen - 1] = 0;
        } else {
            std::memcpy(lpMultiByteStr, str.data(), resLen);
        }
        return resLen;
    } catch (...) {
        return 0;
    }
}

unsigned long GetModuleFileNameW([[maybe_unused]] void* hModule, wchar_t* lpFilename, [[maybe_unused]] unsigned long nSize) {
    if (nSize > 0 && lpFilename) lpFilename[0] = 0;
    return 0;
}

void* GetModuleHandle([[maybe_unused]] const char* lpModuleName) { return nullptr; }
void* GetProcAddress([[maybe_unused]] void* hModule, [[maybe_unused]] const char* lpProcName) { return nullptr; }

#endif
