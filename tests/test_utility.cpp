#include <cassert>
#include <iostream>
#include <string>
#include "types/utility.h"
#include "types/mem.h"

#include <windows.h>
HINSTANCE g_hModule = NULL;

void test_PascalToWString() {
    // Test 1: Null pointer
    std::wstring s1 = PascalToWString(nullptr);
    assert(s1.empty());

    // Test 2: Empty string (length 0)
    wchar_t emptyStr[] = { 0, 0 }; // Length 0, terminated
    std::wstring s2 = PascalToWString(emptyStr);
    assert(s2.empty());

    // Test 3: Normal string
    // "Hello" -> Length 5
    wchar_t helloStr[] = { 5, 'H', 'e', 'l', 'l', 'o', 0 };
    std::wstring s3 = PascalToWString(helloStr);
    assert(s3 == L"Hello");
    assert(s3.length() == 5);

    // Test 4: String with spaces
    wchar_t spaceStr[] = { 3, 'A', ' ', 'B', 0 };
    std::wstring s4 = PascalToWString(spaceStr);
    assert(s4 == L"A B");
    assert(s4.length() == 3);

    // Test 5: Verify it respects length prefix not null terminator (though standard C++ wstring constructor stops at length)
    // Actually std::wstring(ptr, len) takes len chars.
    wchar_t embeddedNull[] = { 3, 'A', 0, 'B', 0 };
    std::wstring s5 = PascalToWString(embeddedNull);
    assert(s5.length() == 3);
    assert(s5[1] == 0);
    assert(s5[2] == 'B');

    std::cout << "PascalToWString tests passed!" << std::endl;
}

void test_IsDateLikeFormat() {
    // TRUE: format already displays a date/time -> caller will SKIP reformatting
    assert(IsDateLikeFormat(L"yyyy-mm-dd"));
    assert(IsDateLikeFormat(L"m/d/yyyy"));
    assert(IsDateLikeFormat(L"h:mm:ss"));
    assert(IsDateLikeFormat(L"[$-409]m/d/yy h:mm AM/PM"));
    assert(IsDateLikeFormat(L"dd mmm yyyy"));

    // FALSE: not a date/time format
    assert(!IsDateLikeFormat(L"General"));
    assert(!IsDateLikeFormat(L"0.00"));
    assert(!IsDateLikeFormat(L"#,##0"));
    assert(!IsDateLikeFormat(L"\"day\" 0")); // only 'd' is inside a quoted literal
    assert(!IsDateLikeFormat(L""));          // empty == General

    std::cout << "IsDateLikeFormat tests passed!" << std::endl;
}

int main() {
    try {
        test_PascalToWString();
        test_IsDateLikeFormat();
    } catch (const std::exception& e) {
        std::cerr << "Test failed with exception: " << e.what() << std::endl;
        return 1;
    }
    return 0;
}
