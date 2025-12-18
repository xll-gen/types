#include <iostream>
#include <cassert>
#include <vector>
#include "types/converters.h"
#include "types/mem.h"
#include "types/protocol_generated.h"
#include "types/xlcall.h"

// xlAutoFree12 is declared in types/mem.h

void TestRangeOverflow() {
    std::cout << "Testing Range Reference Overflow (Security)..." << std::endl;
    flatbuffers::FlatBufferBuilder builder;

    // Create 70,000 rects.
    // This exceeds WORD (65535) count in XLMREF12.
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

    // Expectation: The conversion should identify that 70,000 refs cannot fit in XLMREF12 (max 65535)
    // and return an Error XLOPER.

    LPXLOPER12 op = AnyToXLOPER12(root);

    if (op) {
        if ((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeErr) {
             std::cout << "Success: Returned Error XLOPER as expected." << std::endl;
        } else if ((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeRef) {
             std::cout << "FAILURE: Returned Ref XLOPER." << std::endl;
             std::cout << "Count in XLMREF: " << op->val.mref.lpmref->count << std::endl;
             if (op->val.mref.lpmref->count != num_refs) {
                 std::cout << "CRITICAL: Count was truncated! Expected " << num_refs << ", got " << op->val.mref.lpmref->count << std::endl;
             }
             // Fail the test
             exit(1);
        } else {
             std::cout << "FAILURE: Returned unknown XLOPER type " << op->xltype << std::endl;
             exit(1);
        }

        xlAutoFree12(op);
    } else {
        // NULL return is arguably safe but we prefer Error for feedback.
        std::cout << "Warning: Returned NULL. This prevents overflow but is not ideal." << std::endl;
        // For now, accept NULL as safe.
    }
}

int main() {
    TestRangeOverflow();
    return 0;
}
