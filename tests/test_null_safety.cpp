#include <iostream>
#include <cassert>
#include <vector>
#include "types/converters.h"
#include "types/mem.h"

#ifdef _WIN32
#include <windows.h>
HINSTANCE g_hModule = NULL;
#endif

// We need to link against the library which should provide symbols.

int main() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeMulti;
    op.val.array.rows = 2;
    op.val.array.columns = 2;
    op.val.array.lparray = nullptr; // Malformed!

    std::cout << "Attempting to convert malformed grid..." << std::endl;

    // Test ConvertGrid
    try {
        auto offset = ConvertGrid(&op, builder);
        std::cout << "ConvertGrid: Did not crash! Offset created." << std::endl;

        // Check if it returned an empty grid (which is what we expect after fix)
        // Before fix, we expect crash before this line.
        builder.Finish(offset);
        auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());

        if (grid->rows() == 0 && grid->cols() == 0) {
            std::cout << "ConvertGrid: Returned empty grid (Safe)." << std::endl;
        } else {
            std::cout << "ConvertGrid: Returned non-empty grid? Rows: " << grid->rows() << std::endl;
        }

    } catch (...) {
        std::cout << "ConvertGrid: Caught exception" << std::endl;
    }

    // Test ConvertMultiToAny
    try {
        builder.Clear();
        auto offset = ConvertMultiToAny(op, builder);
        std::cout << "ConvertMultiToAny: Did not crash! Offset created." << std::endl;
    } catch (...) {
         std::cout << "ConvertMultiToAny: Caught exception" << std::endl;
    }

    return 0;
}
