#include <iostream>
#include <string>
#include <vector>
#include <limits>
#include "types/utility.h"

// Basic test for WideToUtf8
void TestWideToUtf8_Basic() {
    std::cout << "TestWideToUtf8_Basic: ";
    std::wstring input = L"Hello World";
    std::string expected = "Hello World";
    std::string output = WideToUtf8(input);

    if (output == expected) {
        std::cout << "PASS" << std::endl;
    } else {
        std::cout << "FAIL (Expected '" << expected << "', got '" << output << "')" << std::endl;
        exit(1);
    }
}

// Test with empty string
void TestWideToUtf8_Empty() {
    std::cout << "TestWideToUtf8_Empty: ";
    std::wstring input = L"";
    std::string output = WideToUtf8(input);

    if (output.empty()) {
        std::cout << "PASS" << std::endl;
    } else {
        std::cout << "FAIL (Expected empty string)" << std::endl;
        exit(1);
    }
}

int main() {
    TestWideToUtf8_Basic();
    TestWideToUtf8_Empty();
    return 0;
}
