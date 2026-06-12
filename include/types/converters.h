#pragma once
#ifdef _WIN32
#include <windows.h>
#else
#include "types/win_compat.h"
#endif
#include "types/xlcall.h"
#include "types/protocol_generated.h" // Needed for protocol:: types
#include <flatbuffers/flatbuffers.h>

// Excel -> Flatbuffers

/**
 * @brief Convert a reference XLOPER12 (xltypeRef / xltypeSRef) into a
 *        protocol::Range, populating its rectangles, format string, and the
 *        owning worksheet name.
 *
 * The `sheet_name` field is resolved via the thread-safe-callable C-API
 * entry points `xlSheetNm` (and, for same-sheet `xltypeSRef` inputs, the
 * no-arg `xlSheetId`). These are legal from non-macro, '$'-registered
 * worksheet functions, matching xll-gen v0.5.0's caller-aware defaults.
 *
 * @note The sheet-name lookup degrades gracefully: if Excel is not callable
 *       (e.g. unit tests with no `MdCallBack12` bound, or a call made outside
 *       a calc context) the field is left empty. The function never throws and
 *       never crashes. Populating `sheet_name` is purely additive — callers
 *       that previously saw an always-empty name are unaffected.
 *
 * @param op      Reference XLOPER12. Other xltypes yield an empty Range
 *                (only the format string is carried).
 * @param builder Destination FlatBufferBuilder.
 * @param format  Optional number-format string for the range.
 * @return Offset to the constructed protocol::Range.
 */
flatbuffers::Offset<protocol::Range> ConvertRange(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder, const std::string& format = "");
flatbuffers::Offset<protocol::Scalar> ConvertScalar(const XLOPER12& cell, flatbuffers::FlatBufferBuilder& builder);
flatbuffers::Offset<protocol::Grid> ConvertGrid(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder);
flatbuffers::Offset<protocol::NumGrid> ConvertNumGrid(FP12* fp, flatbuffers::FlatBufferBuilder& builder);
flatbuffers::Offset<protocol::Any> ConvertAny(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder);

// Flatbuffers -> Excel
LPXLOPER12 AnyToXLOPER12(const protocol::Any* any);
LPXLOPER12 RangeToXLOPER12(const protocol::Range* range);
LPXLOPER12 GridToXLOPER12(const protocol::Grid* grid);
FP12* NumGridToFP12(const protocol::NumGrid* grid);

// Helper for internal use (also exported if needed)
flatbuffers::Offset<protocol::Any> ConvertMultiToAny(const XLOPER12& xMulti, flatbuffers::FlatBufferBuilder& builder);
