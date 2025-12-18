#include "types/converters.h"
#include "types/mem.h"
#include "types/utility.h"
#include "types/PascalString.h"
#include <vector>
#include <algorithm>
#include <limits>

// Excel -> FlatBuffers Converters

flatbuffers::Offset<protocol::Scalar> ConvertScalar(const XLOPER12& cell, flatbuffers::FlatBufferBuilder& builder) {
    if (cell.xltype == xltypeNum) {
        return protocol::CreateScalar(builder, protocol::ScalarValue::Num, protocol::CreateNum(builder, cell.val.num).Union());
    } else if (cell.xltype == xltypeInt) {
        return protocol::CreateScalar(builder, protocol::ScalarValue::Int, protocol::CreateInt(builder, cell.val.w).Union());
    } else if (cell.xltype == xltypeBool) {
        return protocol::CreateScalar(builder, protocol::ScalarValue::Bool, protocol::CreateBool(builder, cell.val.xbool).Union());
    } else if (cell.xltype == xltypeStr) {
         return protocol::CreateScalar(builder, protocol::ScalarValue::Str, protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(cell.val.str))).Union());
    } else if (cell.xltype == xltypeErr) {
         return protocol::CreateScalar(builder, protocol::ScalarValue::Err, protocol::CreateErr(builder, (protocol::XlError)cell.val.err).Union());
    } else {
         return protocol::CreateScalar(builder, protocol::ScalarValue::Nil, protocol::CreateNil(builder).Union());
    }
}

flatbuffers::Offset<protocol::Grid> ConvertGrid(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder) {
    if (op->xltype == xltypeMulti) {
        int rows = op->val.array.rows;
        int cols = op->val.array.columns;
        size_t count = (size_t)rows * cols;

        if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max()) {
             // Return empty grid on overflow or invalid dimensions
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
}

flatbuffers::Offset<protocol::NumGrid> ConvertNumGrid(FP12* fp, flatbuffers::FlatBufferBuilder& builder) {
    if (!fp) return protocol::CreateNumGrid(builder, 0, 0, 0);
    int rows = fp->rows;
    int cols = fp->columns;
    size_t count = (size_t)rows * cols;

    if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max()) {
        return protocol::CreateNumGrid(builder, 0, 0, 0);
    }

    // FP12 array is double[]
    auto vec = builder.CreateVector(fp->array, count);
    return protocol::CreateNumGrid(builder, (uint32_t)rows, (uint32_t)cols, vec);
}

flatbuffers::Offset<protocol::Range> ConvertRange(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder, const std::string& format) {
    auto fmtOff = builder.CreateString(format);

    if (op->xltype & xltypeRef) {
         std::vector<protocol::Rect> rects;
         int count = op->val.mref.lpmref->count;
         for(int i=0; i<count; ++i) {
             auto& r = op->val.mref.lpmref->reftbl[i];
             rects.emplace_back(r.rwFirst, r.rwLast, r.colFirst, r.colLast);
         }
         auto vec = builder.CreateVectorOfStructs(rects);

         return protocol::CreateRange(builder, 0, vec, fmtOff);
    }
    // SRRef
    if (op->xltype & xltypeSRef) {
         std::vector<protocol::Rect> rects;
         auto& r = op->val.sref.ref;
         rects.emplace_back(r.rwFirst, r.rwLast, r.colFirst, r.colLast);
         auto vec = builder.CreateVectorOfStructs(rects);
         return protocol::CreateRange(builder, 0, vec, fmtOff);
    }

    return protocol::CreateRange(builder, 0, 0, fmtOff);
}

// Helper for converting Multi to Any
flatbuffers::Offset<protocol::Any> ConvertMultiToAny(const XLOPER12& op, flatbuffers::FlatBufferBuilder& builder) {
    // Check if it's homogenous numbers -> NumGrid
    // Else -> Grid
    bool allNums = true;
    size_t count = (size_t)op.val.array.rows * op.val.array.columns;

    // Overflow check handled in ConvertGrid calls or here if we use count.
    if (op.val.array.rows < 0 || op.val.array.columns < 0 || count > (size_t)std::numeric_limits<int>::max()) {
        // Fallback to empty Grid
        return protocol::CreateAny(builder, protocol::AnyValue::Grid, protocol::CreateGrid(builder, 0, 0, 0).Union());
    }

    for(size_t i=0; i<count; ++i) {
        if (op.val.array.lparray[i].xltype != xltypeNum) {
            allNums = false;
            break;
        }
    }

    if (allNums) {
         // Create NumGrid
         std::vector<double> nums;
         nums.reserve(count);
         for(size_t i=0; i<count; ++i) nums.push_back(op.val.array.lparray[i].val.num);
         auto vec = builder.CreateVector(nums);
         auto ng = protocol::CreateNumGrid(builder, op.val.array.rows, op.val.array.columns, vec);
         return protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
    } else {
         auto g = ConvertGrid(const_cast<LPXLOPER12>(&op), builder);
         return protocol::CreateAny(builder, protocol::AnyValue::Grid, g.Union());
    }
}

flatbuffers::Offset<protocol::Any> ConvertAny(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder) {
    if (op->xltype == xltypeNum) {
        return protocol::CreateAny(builder, protocol::AnyValue::Num, protocol::CreateNum(builder, op->val.num).Union());
    } else if (op->xltype == xltypeInt) {
        return protocol::CreateAny(builder, protocol::AnyValue::Int, protocol::CreateInt(builder, op->val.w).Union());
    } else if (op->xltype == xltypeBool) {
        return protocol::CreateAny(builder, protocol::AnyValue::Bool, protocol::CreateBool(builder, op->val.xbool).Union());
    } else if (op->xltype == xltypeStr) {
        return protocol::CreateAny(builder, protocol::AnyValue::Str, protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(op->val.str))).Union());
    } else if (op->xltype == xltypeErr) {
        return protocol::CreateAny(builder, protocol::AnyValue::Err, protocol::CreateErr(builder, (protocol::XlError)op->val.err).Union());
    } else if (op->xltype & (xltypeRef | xltypeSRef)) {
        return protocol::CreateAny(builder, protocol::AnyValue::Range, ConvertRange(op, builder).Union());

    } else if (op->xltype & xltypeMulti) {
        return ConvertMultiToAny(*op, builder);
    } else if (op->xltype & (xltypeMissing | xltypeNil)) {
         return protocol::CreateAny(builder, protocol::AnyValue::Nil, protocol::CreateNil(builder).Union());
    }

    return protocol::CreateAny(builder, protocol::AnyValue::Nil, protocol::CreateNil(builder).Union());
}

// FlatBuffers -> Excel Converters

LPXLOPER12 AnyToXLOPER12(const protocol::Any* any) {
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
             op->val.err = (int)any->val_as_Err()->val();
             return op;
        }
        case protocol::AnyValue::Grid: {
             return GridToXLOPER12(any->val_as_Grid());
        }
        case protocol::AnyValue::NumGrid: {
             const protocol::NumGrid* ng = any->val_as_NumGrid();
             int rows = ng->rows();
             int cols = ng->cols();
             size_t count = (size_t)rows * cols;

             if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max() ||
                 !ng->data() || ng->data()->size() < count) {
                 LPXLOPER12 op = NewXLOPER12();
                 op->xltype = xltypeErr | xlbitDLLFree;
                 op->val.err = xlerrValue;
                 return op;
             }

             LPXLOPER12 op = NewXLOPER12();
             op->xltype = xltypeMulti | xlbitDLLFree;
             op->val.array.rows = rows;
             op->val.array.columns = cols;
             op->val.array.lparray = new XLOPER12[count];

             auto data = ng->data();
             for(size_t i=0; i<count; ++i) {
                 op->val.array.lparray[i].xltype = xltypeNum;
                 op->val.array.lparray[i].val.num = data->Get((flatbuffers::uoffset_t)i);
             }
             return op;
        }
        case protocol::AnyValue::Range: {
            return RangeToXLOPER12(any->val_as_Range());
        }
        case protocol::AnyValue::Nil:
        default: {
             LPXLOPER12 op = NewXLOPER12();
             op->xltype = xltypeNil | xlbitDLLFree;
             return op;
        }
    }
}

LPXLOPER12 RangeToXLOPER12(const protocol::Range* range) {
    if (!range) return NULL;

    // Security check: Ensure we don't overflow the WORD count or cause allocation overflow.
    if (range->refs()->size() > std::numeric_limits<WORD>::max()) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeRef | xlbitDLLFree;
    op->val.mref.lpmref = (LPXLMREF12) new char[sizeof(XLMREF12) + sizeof(XLREF12) * range->refs()->size()];
    op->val.mref.idSheet = 0;

    op->val.mref.lpmref->count = (WORD)range->refs()->size();
    for(size_t i=0; i<range->refs()->size(); ++i) {
        const auto* r = range->refs()->Get(i);
        op->val.mref.lpmref->reftbl[i].rwFirst = r->row_first();
        op->val.mref.lpmref->reftbl[i].rwLast = r->row_last();
        op->val.mref.lpmref->reftbl[i].colFirst = r->col_first();
        op->val.mref.lpmref->reftbl[i].colLast = r->col_last();
    }

    return op;
}

LPXLOPER12 GridToXLOPER12(const protocol::Grid* grid) {
    if (!grid) return NULL;

    int rows = grid->rows();
    int cols = grid->cols();
    size_t count = (size_t)rows * cols;

    if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max() ||
        !grid->data() || grid->data()->size() < count) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeMulti | xlbitDLLFree;
    op->val.array.rows = rows;
    op->val.array.columns = cols;
    op->val.array.lparray = new XLOPER12[count];

    for (size_t i = 0; i < count; ++i) {
        auto scalar = grid->data()->Get((flatbuffers::uoffset_t)i);
        auto& cell = op->val.array.lparray[i];
        cell.xltype = xltypeNil; // Default

        switch(scalar->val_type()) {
            case protocol::ScalarValue::Num:
                cell.xltype = xltypeNum;
                cell.val.num = scalar->val_as_Num()->val();
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
                cell.xltype = xltypeStr;
                std::wstring ws = StringToWString(scalar->val_as_Str()->val()->str());
                auto vec = WStringToPascalString(ws);
                cell.val.str = new XCHAR[vec.size()];
                std::copy(vec.begin(), vec.end(), cell.val.str);
                break;
            }
            case protocol::ScalarValue::Err:
                cell.xltype = xltypeErr;
                cell.val.err = (int)scalar->val_as_Err()->val();
                break;
            default:
                break;
        }
    }

    return op;
}

FP12* NumGridToFP12(const protocol::NumGrid* grid) {
    if (!grid) return NULL;
    int rows = grid->rows();
    int cols = grid->cols();
    size_t count = (size_t)rows * cols;

    if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max() ||
        !grid->data() || grid->data()->size() < count) {
        // Return 0x0
        return NewFP12(0, 0);
    }

    FP12* fp = NewFP12(rows, cols);

    const auto* data = grid->data();
    for(size_t i=0; i<count; ++i) {
        fp->array[i] = data->Get((flatbuffers::uoffset_t)i);
    }

    return fp;
}
