#include <iostream>
#include <cassert>
#include <vector>
#include <string>
#include <limits>
#include "types/converters.h"
#include "types/utility.h"
#include "types/mem.h"

#ifdef _WIN32
HINSTANCE g_hModule = NULL;
#endif

extern "C" void __stdcall xlAutoFree12(LPXLOPER12 p);

void TestDoSProtection() {
    std::cout << "Testing DoS Protection..." << std::endl;
    // Create a very large string > 10MB
    size_t largeSize = 11000000;
    std::string largeStr(largeSize, 'A');

    try {
        std::wstring ws = StringToWString(largeStr);
        // If we implement throwing:
        std::cerr << "Test failed: StringToWString should have thrown for 11MB string" << std::endl;
        exit(1);
    } catch (...) {
        std::cout << "Caught expected exception for huge string." << std::endl;
    }
}

void TestTruncationFix() {
    std::cout << "Testing Truncation Fix..." << std::endl;
    // Create a string > 32767 chars
    size_t len = 40000;
    std::string longStr(len, 'B');

    flatbuffers::FlatBufferBuilder builder;
    std::vector<flatbuffers::Offset<protocol::Scalar>> elements;
    elements.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Str,
        protocol::CreateStr(builder, builder.CreateString(longStr)).Union()));

    auto vec = builder.CreateVector(elements);
    auto grid = protocol::CreateGrid(builder, 1, 1, vec);
    builder.Finish(grid);

    auto* root = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 res = GridToXLOPER12(root);

    assert(res->xltype == (xltypeMulti | xlbitDLLFree));
    assert(res->val.array.rows == 1);
    assert(res->val.array.columns == 1);

    LPXLOPER12 cell = &res->val.array.lparray[0];
    assert(cell->xltype == xltypeStr);

    // Check length is truncated to 32767
    assert(cell->val.str[0] == 32767);

    // Check content is valid (not garbage)
    // We expect 'B's.
    // Check the last char (index 32767)
    assert(cell->val.str[32767] == L'B');
    // Check first char
    assert(cell->val.str[1] == L'B');

    xlAutoFree12(res);
    std::cout << "TestTruncationFix passed" << std::endl;
}

int main() {
    TestDoSProtection();
    TestTruncationFix();
    return 0;
}
