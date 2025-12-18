#include <iostream>
#include <vector>
#include <string>
#include <limits>
#include "types/converters.h"
#include "types/mem.h"
#include "types/protocol_generated.h"
#include "types/xlcall.h"

// Explicit declaration of xlAutoFree12 as it is used for cleanup
// (It is declared in types/mem.h, so we just include it)

void TestRangeOverflow() {
    std::cout << "Testing Range Reference Overflow (Regression Check)..." << std::endl;
    flatbuffers::FlatBufferBuilder builder;

    // Create 70,000 rects.
    size_t num_refs = 70000;
    std::vector<protocol::Rect> rects;
    rects.reserve(num_refs);
    for(size_t i=0; i<num_refs; ++i) {
        rects.emplace_back(0, 1, 0, 1);
    }

    auto refs_vec = builder.CreateVectorOfStructs(rects);
    auto fmt = builder.CreateString("");
    auto range = protocol::CreateRange(builder, 0, refs_vec, fmt);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Range, range.Union());
    builder.Finish(any);

    const protocol::Any* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());

    LPXLOPER12 op = AnyToXLOPER12(root);

    if (op) {
        if ((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeErr) {
             std::cout << "Success: Returned Error XLOPER as expected." << std::endl;
        } else {
             std::cout << "FAILURE: Returned non-error XLOPER type " << op->xltype << std::endl;
             exit(1);
        }
        xlAutoFree12(op);
    } else {
        std::cout << "Warning: Returned NULL." << std::endl;
    }
}

void TestGridStringOversize() {
    std::cout << "Testing Grid String Oversize Handling..." << std::endl;
    flatbuffers::FlatBufferBuilder builder;

    // Create a string longer than 32767 chars (Excel limit)
    // 32768 chars 'A'
    std::string large_str(32768, 'A');

    auto str_off = builder.CreateString(large_str);
    auto scalar = protocol::CreateScalar(builder, protocol::ScalarValue::Str, protocol::CreateStr(builder, str_off).Union());

    std::vector<flatbuffers::Offset<protocol::Scalar>> elements;
    elements.push_back(scalar);
    auto vec = builder.CreateVector(elements);

    // Create 1x1 Grid
    auto grid = protocol::CreateGrid(builder, 1, 1, vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Grid, grid.Union());
    builder.Finish(any);

    const protocol::Any* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());

    LPXLOPER12 op = AnyToXLOPER12(root);

    if (op) {
        // It should be a Multi containing a Str
        if ((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeMulti) {
             LPXLOPER12 cell = op->val.array.lparray; // 0,0
             if (cell && (cell->xltype == xltypeStr)) {
                 int len = (int)cell->val.str[0];
                 std::cout << "String length returned: " << len << std::endl;
                 // With strict clamping and MultiByteToWideChar behavior,
                 // requesting conversion of 32768 chars into 32767 buffer fails.
                 // So we expect 0 (or empty string).
                 // Alternatively, if the implementation truncates properly (complex), it might be 32767.
                 // But our plan is to allow failure (safe empty string).
                 if (len == 0) {
                     std::cout << "Success: String truncated/empty due to oversize." << std::endl;
                 } else {
                     std::cout << "Info: String returned with length " << len << std::endl;
                     if (len > 32767) {
                         std::cout << "FAILURE: Length exceeds 32767." << std::endl;
                         exit(1);
                     }
                 }
             } else {
                 std::cout << "FAILURE: Grid element is not string." << std::endl;
                 exit(1);
             }
        } else {
             std::cout << "FAILURE: Returned non-Multi XLOPER " << op->xltype << std::endl;
             exit(1);
        }
        xlAutoFree12(op);
    } else {
        std::cout << "FAILURE: Returned NULL." << std::endl;
        exit(1);
    }
}

int main() {
    TestRangeOverflow();
    TestGridStringOversize();
    return 0;
}
