#include "types/converters.h"
#include "types/mem.h"
#include "types/utility.h"
#include "types/ScopeGuard.h"
#include "types/ScopedXLOPER12.h"
#include <vector>
#include <algorithm>
#include <limits>
#include <new>
#include <cstring> // for std::memset


// Excel -> FlatBuffers Converters

// Helper functions (internal)
static int ProtocolErrorToExcel(protocol::XlError err) {
    return (int)err - 2000;
}

static protocol::XlError ExcelErrorToProtocol(int err) {
    return (protocol::XlError)(err + 2000);
}

// Ownership bits that may ride on xltype without changing the logical type.
// xlbitDLLFree marks memory we allocated (freed via xlAutoFree12), xlbitXLFree
// marks Excel-owned memory (freed via xlFree). Every reader that switches on
// the type of an XLOPER12 — top-level or xltypeMulti element — must mask
// these bits first.
static const DWORD kXlOwnerBits = (DWORD)(xlbitDLLFree | xlbitXLFree);

static inline DWORD BaseXlType(const XLOPER12& op) {
    return op.xltype & ~kXlOwnerBits;
}

// The canonical "return an error XLOPER12 to Excel" block (R5 §4.7 dedup).
// This 4-line pattern used to be copy-pasted ~10x; security fixes
// (BUG-014/015/017 lineage) must land here exactly once.
static LPXLOPER12 MakeErrXLOPER12(int err) {
    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeErr | xlbitDLLFree;
    op->val.err = err;
    return op;
}

// Shared grid-dimension validation (R5 §4.7 dedup, BUG-015/017 history).
// Rejects negative dimensions, a rows*cols product that would overflow
// size_t (relevant on 32-bit x86 builds), and element counts above INT_MAX
// (Excel's practical limit and our downstream loop assumption). On success
// *outCount holds rows*cols. Do not weaken any individual check.
static bool ValidateGridDims(int rows, int cols, size_t* outCount) {
    *outCount = 0;
    if (rows < 0 || cols < 0) return false;
    if (cols > 0 && (size_t)rows > SIZE_MAX / (size_t)cols) return false;
    size_t count = (size_t)rows * (size_t)cols;
    if (count > (size_t)std::numeric_limits<int>::max()) return false;
    *outCount = count;
    return true;
}

flatbuffers::Offset<protocol::Scalar> ConvertScalar(const XLOPER12& cell, flatbuffers::FlatBufferBuilder& builder) {
    try {
        // Mask xlbitDLLFree/xlbitXLFree: multi elements built by
        // GridToXLOPER12 carry xlbitDLLFree on their string cells, and this
        // function must still classify them correctly when such an array is
        // converted back (round-trip / echo paths).
        const DWORD type = BaseXlType(cell);
        if (type == xltypeNum) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Num, protocol::CreateNum(builder, cell.val.num).Union());
        } else if (type == xltypeInt) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Int, protocol::CreateInt(builder, cell.val.w).Union());
        } else if (type == xltypeBool) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Bool, protocol::CreateBool(builder, cell.val.xbool).Union());
        } else if (type == xltypeStr) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Str, protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(cell.val.str))).Union());
        } else if (type == xltypeErr) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Err, protocol::CreateErr(builder, ExcelErrorToProtocol(cell.val.err)).Union());
        } else {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Nil, protocol::CreateNil(builder).Union());
        }
    } catch (...) {
        return protocol::CreateScalar(builder, protocol::ScalarValue::Nil, protocol::CreateNil(builder).Union());
    }
}

flatbuffers::Offset<protocol::Grid> ConvertGrid(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder) {
    try {
        if (BaseXlType(*op) == xltypeMulti) {
            int rows = op->val.array.rows;
            int cols = op->val.array.columns;

            size_t count = 0;
            if (!ValidateGridDims(rows, cols, &count)) {
                // Return empty grid on overflow or invalid dimensions
                return protocol::CreateGrid(builder, 0, 0, 0);
            }

            // Safety check: if grid dimensions are non-zero, lparray must not be null
            if (count > 0 && !op->val.array.lparray) {
                return protocol::CreateGrid(builder, 0, 0, 0);
            }

            std::vector<flatbuffers::Offset<protocol::Scalar>> elements;
            elements.reserve(count);

            for (size_t i = 0; i < count; ++i) {
                elements.push_back(ConvertScalar(op->val.array.lparray[i], builder));
            }

            auto vec = builder.CreateVector(elements);
            return protocol::CreateGrid(builder, (uint32_t)rows, (uint32_t)cols, vec);
        }

        // Handle scalar as 1x1 Grid
        std::vector<flatbuffers::Offset<protocol::Scalar>> elements;
        elements.push_back(ConvertScalar(*op, builder));

        auto vec = builder.CreateVector(elements);
        return protocol::CreateGrid(builder, 1, 1, vec);
    } catch (...) {
        return protocol::CreateGrid(builder, 0, 0, 0);
    }
}

flatbuffers::Offset<protocol::NumGrid> ConvertNumGrid(FP12* fp, flatbuffers::FlatBufferBuilder& builder) {
    try {
        if (!fp) return protocol::CreateNumGrid(builder, 0, 0, 0);
        int rows = fp->rows;
        int cols = fp->columns;

        size_t count = 0;
        if (!ValidateGridDims(rows, cols, &count)) {
            return protocol::CreateNumGrid(builder, 0, 0, 0);
        }

        // Safety check for vector allocation size (though FlatBuffers handles it, explicit check prevents huge alloc attempts)
        if (count > SIZE_MAX / sizeof(double)) {
            return protocol::CreateNumGrid(builder, 0, 0, 0);
        }

        // FP12 array is double[]
        auto vec = builder.CreateVector(fp->array, count);
        return protocol::CreateNumGrid(builder, (uint32_t)rows, (uint32_t)cols, vec);
    } catch (...) {
        return protocol::CreateNumGrid(builder, 0, 0, 0);
    }
}

// Resolve the workbook/worksheet name ("[Book1]Sheet1") for a reference
// XLOPER12 using only C-API entry points that are legal from non-macro,
// thread-safe ('$'-registered) worksheet functions. xll-gen v0.5.0 makes
// caller-aware functions non-macro by default, so xlSheetNm/xlSheetId — both
// in the documented thread-safe-callable set — are the only sanctioned path.
//
// Mechanism, by reference flavour:
//   - xltypeRef  (external reference, carries idSheet): pass `op` straight to
//     xlSheetNm, which reads val.mref.idSheet and returns the name string.
//   - xltypeSRef (same-sheet reference, no idSheet): first call the no-arg
//     xlSheetId to learn the active sheet id, then hand that ref to xlSheetNm.
//
// Returns the name as UTF-8. On ANY failure — no live Excel (pexcel12 NULL in
// unit tests → xlretFailed), not in a calc context, wrong return type, or an
// exception in conversion — it degrades to an empty string. It never throws.
// The xlSheetNm result is an Excel-allocated xltypeStr owned by Excel; the
// ScopedXLOPER12Result wrapper releases it with xlFree on scope exit.
static std::string LookupSheetName(LPXLOPER12 op) {
    try {
        if (!op) return std::string();

        const DWORD baseType = BaseXlType(*op);

        // The xltypeRef path can use op directly (it carries idSheet). The
        // xltypeSRef path needs a ref-with-idSheet, which xlSheetId supplies.
        ScopedXLOPER12Result xSheetId; // holds the xlSheetId result (xltypeRef)
        LPXLOPER12 refForName = nullptr;

        if (baseType == xltypeRef) {
            refForName = op;
        } else if (baseType == xltypeSRef) {
            // No-arg xlSheetId returns an xltypeRef for the active sheet.
            if (Excel12(xlSheetId, xSheetId, 0) != xlretSuccess) {
                return std::string();
            }
            if (BaseXlType(*xSheetId.get()) != xltypeRef) {
                return std::string();
            }
            refForName = xSheetId.get();
        } else {
            return std::string();
        }

        ScopedXLOPER12Result xName;
        if (Excel12(xlSheetNm, xName, 1, refForName) != xlretSuccess) {
            return std::string();
        }
        if (BaseXlType(*xName.get()) != xltypeStr || !xName.get()->val.str) {
            return std::string();
        }

        // Excel returns a Pascal-style wide string ("[Book1]Sheet1").
        std::wstring ws = PascalToWString(xName.get()->val.str);
        return WideToUtf8(ws);
    } catch (...) {
        return std::string();
    }
    // ScopedXLOPER12Result destructors free any Excel-allocated payloads here.
}

flatbuffers::Offset<protocol::Range> ConvertRange(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder, const std::string& format) {
    try {
        // Resolve the sheet name once per call (not per rect). Empty on any
        // failure / outside a live Excel calc context — purely additive, so
        // existing consumers that never saw a populated name are unaffected.
        std::string sheetName = LookupSheetName(op);
        flatbuffers::Offset<flatbuffers::String> sheetOff =
            sheetName.empty() ? 0 : builder.CreateString(sheetName);

        auto fmtOff = builder.CreateString(format);

        if (op->xltype & xltypeRef) {
            std::vector<protocol::Rect> rects;
            int count = op->val.mref.lpmref->count;
            for(int i=0; i<count; ++i) {
                auto& r = op->val.mref.lpmref->reftbl[i];
                rects.emplace_back(r.rwFirst, r.rwLast, r.colFirst, r.colLast);
            }
            auto vec = builder.CreateVectorOfStructs(rects);

            return protocol::CreateRange(builder, sheetOff, vec, fmtOff);
        }
        // SRRef
        if (op->xltype & xltypeSRef) {
            std::vector<protocol::Rect> rects;
            auto& r = op->val.sref.ref;
            rects.emplace_back(r.rwFirst, r.rwLast, r.colFirst, r.colLast);
            auto vec = builder.CreateVectorOfStructs(rects);
            return protocol::CreateRange(builder, sheetOff, vec, fmtOff);
        }

        return protocol::CreateRange(builder, sheetOff, 0, fmtOff);
    } catch (...) {
        return protocol::CreateRange(builder, 0, 0, 0);
    }
}

// Helper for converting Multi to Any
flatbuffers::Offset<protocol::Any> ConvertMultiToAny(const XLOPER12& op, flatbuffers::FlatBufferBuilder& builder) {
    try {
        // Check if it's homogenous numbers -> NumGrid
        // Else -> Grid
        bool allNums = true;

        size_t count = 0;
        if (!ValidateGridDims(op.val.array.rows, op.val.array.columns, &count)) {
            // Fallback to empty Grid
            return protocol::CreateAny(builder, protocol::AnyValue::Grid, protocol::CreateGrid(builder, 0, 0, 0).Union());
        }

        // Safety check: if grid dimensions are non-zero, lparray must not be null
        if (count > 0 && !op.val.array.lparray) {
             return protocol::CreateAny(builder, protocol::AnyValue::Grid, protocol::CreateGrid(builder, 0, 0, 0).Union());
        }

        for (size_t i = 0; i < count; ++i) {
            // Mask ownership bits — see ConvertScalar.
            if (BaseXlType(op.val.array.lparray[i]) != xltypeNum) {
                allNums = false;
                break;
            }
        }

        if (allNums) {
            // Create NumGrid
            // Optimization: Use CreateUninitializedVector to write directly to FlatBuffer memory,
            // avoiding intermediate std::vector allocation and redundant copy.
            double* buf = nullptr;
            auto vec = builder.CreateUninitializedVector<double>(count, &buf);
            for (size_t i = 0; i < count; ++i) {
                buf[i] = op.val.array.lparray[i].val.num;
            }
            auto ng = protocol::CreateNumGrid(builder, op.val.array.rows, op.val.array.columns, vec);
            return protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
        } else {
            auto g = ConvertGrid(const_cast<LPXLOPER12>(&op), builder);
            return protocol::CreateAny(builder, protocol::AnyValue::Grid, g.Union());
        }
    } catch (...) {
        return protocol::CreateAny(builder, protocol::AnyValue::Err,
                                   protocol::CreateErr(builder, protocol::XlError::Unknown).Union());
    }
}

flatbuffers::Offset<protocol::Any> ConvertAny(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder) {
    try {
        // Mask ownership bits so XLOPER12s produced by our own
        // AnyToXLOPER12/GridToXLOPER12 (which carry xlbitDLLFree) classify
        // correctly when fed back through the Excel->FlatBuffers direction.
        const DWORD type = BaseXlType(*op);
        if (type == xltypeNum) {
            return protocol::CreateAny(builder, protocol::AnyValue::Num,
                                       protocol::CreateNum(builder, op->val.num).Union());
        } else if (type == xltypeInt) {
            return protocol::CreateAny(builder, protocol::AnyValue::Int,
                                       protocol::CreateInt(builder, op->val.w).Union());
        } else if (type == xltypeBool) {
            return protocol::CreateAny(builder, protocol::AnyValue::Bool,
                                       protocol::CreateBool(builder, op->val.xbool).Union());
        } else if (type == xltypeStr) {
            return protocol::CreateAny(builder, protocol::AnyValue::Str,
                                       protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(op->val.str)))
                                           .Union());
        } else if (type == xltypeErr) {
            return protocol::CreateAny(builder, protocol::AnyValue::Err,
                                       protocol::CreateErr(builder, ExcelErrorToProtocol(op->val.err)).Union());
        } else if (type & (xltypeRef | xltypeSRef)) {
            return protocol::CreateAny(builder, protocol::AnyValue::Range, ConvertRange(op, builder).Union());

        } else if (type & xltypeMulti) {
            return ConvertMultiToAny(*op, builder);
        } else if (type & (xltypeMissing | xltypeNil)) {
            return protocol::CreateAny(builder, protocol::AnyValue::Nil, protocol::CreateNil(builder).Union());
        }

        return protocol::CreateAny(builder, protocol::AnyValue::Nil, protocol::CreateNil(builder).Union());
    } catch (...) {
        return protocol::CreateAny(builder, protocol::AnyValue::Err,
                                   protocol::CreateErr(builder, protocol::XlError::Unknown).Union());
    }
}

// FlatBuffers -> Excel Converters

LPXLOPER12 AnyToXLOPER12(const protocol::Any* any) {
    try {
        if (!any) {
            LPXLOPER12 op = NewXLOPER12();
            op->xltype = xltypeNil | xlbitDLLFree;
            return op;
        }

        switch (any->val_type()) {
            case protocol::AnyValue::Num: {
                LPXLOPER12 op = NewXLOPER12();
                op->xltype = xltypeNum | xlbitDLLFree;
                op->val.num = any->val_as_Num()->val();
                return op;
            }
            case protocol::AnyValue::Date: {
                LPXLOPER12 op = NewXLOPER12();
                op->xltype = xltypeNum | xlbitDLLFree;
                op->val.num = any->val_as_Date()->serial();
                return op;
            }
            case protocol::AnyValue::Int: {
                LPXLOPER12 op = NewXLOPER12();
                op->xltype = xltypeInt | xlbitDLLFree;
                op->val.w = any->val_as_Int()->val();
                return op;
            }
            case protocol::AnyValue::Bool: {
                LPXLOPER12 op = NewXLOPER12();
                op->xltype = xltypeBool | xlbitDLLFree;
                op->val.xbool = any->val_as_Bool()->val();
                return op;
            }
            case protocol::AnyValue::Str: {
                std::wstring ws = StringToWString(any->val_as_Str()->val()->str());
                return NewExcelString(ws);
            }
            case protocol::AnyValue::Err: {
                 LPXLOPER12 op = NewXLOPER12();
                 op->xltype = xltypeErr | xlbitDLLFree;
                 op->val.err = ProtocolErrorToExcel(any->val_as_Err()->val());
                 return op;
            }
            case protocol::AnyValue::Grid: {
                 return GridToXLOPER12(any->val_as_Grid());
            }
            case protocol::AnyValue::NumGrid: {
                 const protocol::NumGrid* ng = any->val_as_NumGrid();
                 int rows = ng->rows();
                 int cols = ng->cols();

                 size_t count = 0;
                 if (!ValidateGridDims(rows, cols, &count) ||
                     !ng->data() || ng->data()->size() < count) {
                     return MakeErrXLOPER12(xlerrValue);
                 }

                 // Check for allocation overflow (size_t)
                 if (count > SIZE_MAX / sizeof(XLOPER12)) {
                     return MakeErrXLOPER12(xlerrValue);
                 }

                 LPXLOPER12 op = NewXLOPER12();
                 op->xltype = xltypeMulti | xlbitDLLFree;
                 op->val.array.rows = rows;
                 op->val.array.columns = cols;
                 op->val.array.lparray = nullptr;

                 // BUG-017: Guard must be declared before allocation to protect 'op'
                 ScopeGuard guard([&]() {
                     if (op->val.array.lparray) delete[] op->val.array.lparray;
                     ReleaseXLOPER12(op);
                 });

                 op->val.array.lparray = new XLOPER12[count];

                 auto data = ng->data();
                 for(size_t i=0; i<count; ++i) {
                     op->val.array.lparray[i].xltype = xltypeNum;
                     op->val.array.lparray[i].val.num = data->Get((flatbuffers::uoffset_t)i);
                 }

                 guard.Dismiss();
                 return op;
            }
            case protocol::AnyValue::Range: {
                return RangeToXLOPER12(any->val_as_Range());
            }
            case protocol::AnyValue::RefCache: {
                const auto* rc = any->val_as_RefCache();
                if (rc && rc->key()) {
                    std::wstring ws = StringToWString(rc->key()->str());
                    return NewExcelString(ws);
                }
                return MakeErrXLOPER12(xlerrNA);
            }
            case protocol::AnyValue::AsyncHandle: {
                // Return "#ASYNC!" string to indicate handle
                return NewExcelString(L"#ASYNC!");
            }
            case protocol::AnyValue::Nil:
            default: {
                 LPXLOPER12 op = NewXLOPER12();
                 op->xltype = xltypeNil | xlbitDLLFree;
                 return op;
            }
        }
    } catch (...) {
        return MakeErrXLOPER12(xlerrValue);
    }
}

LPXLOPER12 RangeToXLOPER12(const protocol::Range* range) {
    if (!range) {
        return MakeErrXLOPER12(xlerrValue);
    }

    // Validate refs existence and count limits (XLOPER12 ref count is WORD/16-bit)
    if (!range->refs() || range->refs()->size() > 65535) {
        return MakeErrXLOPER12(xlerrValue);
    }

    try {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeRef | xlbitDLLFree;
        op->val.mref.lpmref = nullptr;

        // Safe allocation size calculation
        size_t refs_count = range->refs()->size();

        // Guard against leaks if new throws
        // BUG-014: Ensure lpmref is freed if set
        ScopeGuard guard([&]() {
            if (op->val.mref.lpmref) {
                 delete[] (char*)op->val.mref.lpmref;
            }
            ReleaseXLOPER12(op);
        });

        // XLMREF12 struct has 1 ref. We need space for (refs_count) refs in total.
        op->val.mref.lpmref = (LPXLMREF12) new char[sizeof(XLMREF12) + sizeof(XLREF12) * refs_count];
        op->val.mref.idSheet = 0;

        op->val.mref.lpmref->count = (WORD)refs_count;
        for(size_t i=0; i<refs_count; ++i) {
            const auto* r = range->refs()->Get(i);
            op->val.mref.lpmref->reftbl[i].rwFirst = r->row_first();
            op->val.mref.lpmref->reftbl[i].rwLast = r->row_last();
            op->val.mref.lpmref->reftbl[i].colFirst = r->col_first();
            op->val.mref.lpmref->reftbl[i].colLast = r->col_last();
        }

        guard.Dismiss();
        return op;
    } catch (...) {
        return MakeErrXLOPER12(xlerrValue);
    }
}

LPXLOPER12 GridToXLOPER12(const protocol::Grid* grid) {
    if (!grid) {
        return MakeErrXLOPER12(xlerrValue);
    }

    int rows = grid->rows();
    int cols = grid->cols();

    size_t count = 0;
    if (!ValidateGridDims(rows, cols, &count) ||
        !grid->data() || grid->data()->size() != count) {
        return MakeErrXLOPER12(xlerrValue);
    }

    // Check for allocation overflow (size_t)
    if (count > SIZE_MAX / sizeof(XLOPER12)) {
         return MakeErrXLOPER12(xlerrValue);
    }

    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeMulti | xlbitDLLFree;
    op->val.array.rows = rows;
    op->val.array.columns = cols;

    // RAII guard to clean up if exception happens or return early.
    // Ownership contract: only element strings carrying xlbitDLLFree were
    // allocated by us — anything else is borrowed and must not be deleted.
    ScopeGuard guard([&]() {
        if (op->val.array.lparray) {
            for(size_t i=0; i<count; ++i) {
                 XLOPER12& elem = op->val.array.lparray[i];
                 if ((elem.xltype & xlbitDLLFree) && BaseXlType(elem) == xltypeStr && elem.val.str) {
                     delete[] elem.val.str;
                 }
            }
            delete[] op->val.array.lparray;
        }
        ReleaseXLOPER12(op);
    });

    try {
        op->val.array.lparray = new XLOPER12[count];
        std::memset(op->val.array.lparray, 0, count * sizeof(XLOPER12));

        for (size_t i = 0; i < count; ++i) {
            auto scalar = grid->data()->Get((flatbuffers::uoffset_t)i);
            auto& cell = op->val.array.lparray[i];
            cell.xltype = xltypeNil; // Default

            switch(scalar->val_type()) {
                case protocol::ScalarValue::Num:
                    cell.xltype = xltypeNum;
                    cell.val.num = scalar->val_as_Num()->val();
                    break;
                case protocol::ScalarValue::Date:
                    cell.xltype = xltypeNum;
                    cell.val.num = scalar->val_as_Date()->serial();
                    break;
                case protocol::ScalarValue::Int:
                    cell.xltype = xltypeInt;
                    cell.val.w = scalar->val_as_Int()->val();
                    break;
                case protocol::ScalarValue::Bool:
                    cell.xltype = xltypeBool;
                    cell.val.xbool = scalar->val_as_Bool()->val();
                    break;
                case protocol::ScalarValue::Str: {
                    // xlbitDLLFree on the element marks the string as
                    // DLL-owned: xlAutoFree12 (and the guard above) only
                    // delete[] element strings carrying this bit. Excel
                    // ignores the bit on inner elements, so it is purely
                    // our ownership marker; our own readers (ConvertScalar,
                    // ConvertMultiToAny, ConvertAny) mask it before type
                    // dispatch.
                    cell.xltype = xltypeStr | xlbitDLLFree;
                    const auto* fbStr = scalar->val_as_Str()->val();
                    const char* utf8 = fbStr ? fbStr->c_str() : nullptr;
                    Utf8ToExcelString(utf8, cell.val.str);
                    break;
                }
                case protocol::ScalarValue::Err:
                    cell.xltype = xltypeErr;
                    cell.val.err = ProtocolErrorToExcel(scalar->val_as_Err()->val());
                    break;
                default:
                    break;
            }
        }
    } catch (...) {
        return MakeErrXLOPER12(xlerrValue);
    }

    guard.Dismiss();
    return op;
}

FP12* NumGridToFP12(const protocol::NumGrid* grid) {
    try {
        if (!grid) return NewFP12(0, 0);
        int rows = grid->rows();
        int cols = grid->cols();

        size_t count = 0;
        if (!ValidateGridDims(rows, cols, &count) ||
            !grid->data() || grid->data()->size() != count) {
            // Return 0x0
            return NewFP12(0, 0);
        }

        // Check for allocation overflow (size_t) for FP12 (8 bytes per double)
        // NewFP12 calculates size: 2*int + count*8.
        if (count > (SIZE_MAX - 2*sizeof(int)) / sizeof(double)) {
             return NewFP12(0, 0);
        }

        FP12* fp = NewFP12(rows, cols);

        const auto* data = grid->data();
        if (data) {
             // Optimization: Use memcpy for bulk copy of doubles
             // FlatBuffers stores vector data contiguously in little-endian.
             // On x86/ARM little-endian systems, this is a direct copy.
             std::memcpy(fp->array, data->data(), count * sizeof(double));
        }

        return fp;
    } catch (...) {
        try {
            return NewFP12(0, 0);
        } catch (...) {
            return nullptr;
        }
    }
}
