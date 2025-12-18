#include <iostream>
#include <cassert>
#include <cstring>
#include <vector>
#include <exception>
#include <limits>

#ifdef _WIN32
#include <windows.h>
HINSTANCE g_hModule = NULL;
#endif

#include "types/converters.h"
#include "types/mem.h"
#include "flatbuffers/flatbuffers.h"
#include "types/protocol_generated.h"
#include "types/xlcall.h"

extern "C" void __stdcall xlAutoFree12(LPXLOPER12 p);

void TestNumGridDataMismatch() {
    std::cout << "Testing NumGrid Data Mismatch..." << std::endl;
    flatbuffers::FlatBufferBuilder builder;

    std::vector<double> empty_data;
    auto data_vec = builder.CreateVector(empty_data);

    auto ng = protocol::CreateNumGrid(builder, 10, 10, data_vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
    builder.Finish(any);

    const protocol::Any* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());

    LPXLOPER12 op = AnyToXLOPER12(root);

    if (op) {
        if (op->xltype == (xltypeErr | xlbitDLLFree)) {
            std::cout << "Success: Returned Error XLOPER as expected." << std::endl;
        } else {
            std::cout << "FAILURE: Returned non-error XLOPER (type=" << op->xltype << ")" << std::endl;
            exit(1);
        }
        xlAutoFree12(op);
    } else {
        std::cout << "FAILURE: Returned nullptr." << std::endl;
        exit(1);
    }
}

void TestNumGridOverflow() {
    std::cout << "Testing NumGrid Overflow..." << std::endl;
    flatbuffers::FlatBufferBuilder builder;

    // rows=46341, cols=46341 -> rows*cols > INT_MAX
    int rows = 46341;
    int cols = 46341;

    std::vector<double> small_data;
    auto data_vec = builder.CreateVector(small_data);

    auto ng = protocol::CreateNumGrid(builder, rows, cols, data_vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
    builder.Finish(any);

    const protocol::Any* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());

    try {
        LPXLOPER12 op = AnyToXLOPER12(root);

        if (op) {
            if (op->xltype == (xltypeErr | xlbitDLLFree)) {
                std::cout << "Success: Returned Error XLOPER as expected." << std::endl;
            } else {
                 // It might be possible that it returns 0x0 size grid if checks fail differently?
                 // But we put check at top.
                std::cout << "FAILURE: Returned non-error XLOPER (type=" << op->xltype << ")" << std::endl;
                exit(1);
            }
            xlAutoFree12(op);
        }
    } catch (const std::exception& e) {
        std::cout << "FAILURE: Caught exception: " << e.what() << std::endl;
        exit(1);
    }
}

int main() {
    TestNumGridDataMismatch();
    std::cout.flush();
    TestNumGridOverflow();

    std::cout << "Repro tests finished successfully." << std::endl;
    return 0;
}
