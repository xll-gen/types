#include "types/converters.h"
#include "types/mem.h"
#include "types/utility.h"
#include "types/PascalString.h"
#include "types/ScopeGuard.h"
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

flatbuffers::Offset<protocol::Scalar> ConvertScalar(const XLOPER12& cell, flatbuffers::FlatBufferBuilder& builder) {
    try {
        if (cell.xltype == xltypeNum) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Num, protocol::CreateNum(builder, cell.val.num).Union());
        } else if (cell.xltype == xltypeInt) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Int, protocol::CreateInt(builder, cell.val.w).Union());
        } else if (cell.xltype == xltypeBool) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Bool, protocol::CreateBool(builder, cell.val.xbool).Union());
        } else if (cell.xltype == xltypeStr) {
            return protocol::CreateScalar(builder, protocol::ScalarValue::Str, protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(cell.val.str))).Union());
        } else if (cell.xltype == xltypeErr) {
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
        if (op->xltype == xltypeMulti) {
            int rows = op->val.array.rows;
            int cols = op->val.array.columns;
            size_t count = (size_t)rows * cols;

            if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max()) {
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
        size_t count = (size_t)rows * cols;

        if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max()) {
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

flatbuffers::Offset<protocol::Range> ConvertRange(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder, const std::string& format) {
    try {
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
        size_t count = (size_t)op.val.array.rows * op.val.array.columns;

        // Overflow check handled in ConvertGrid calls or here if we use count.
        if (op.val.array.rows < 0 || op.val.array.columns < 0 || count > (size_t)std::numeric_limits<int>::max()) {
            // Fallback to empty Grid
            return protocol::CreateAny(builder, protocol::AnyValue::Grid, protocol::CreateGrid(builder, 0, 0, 0).Union());
        }

        // Safety check: if grid dimensions are non-zero, lparray must not be null
        if (count > 0 && !op.val.array.lparray) {
             return protocol::CreateAny(builder, protocol::AnyValue::Grid, protocol::CreateGrid(builder, 0, 0, 0).Union());
        }

        for (size_t i = 0; i < count; ++i) {
            if (op.val.array.lparray[i].xltype != xltypeNum) {
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
        if (op->xltype == xltypeNum) {
            return protocol::CreateAny(builder, protocol::AnyValue::Num,
                                       protocol::CreateNum(builder, op->val.num).Union());
        } else if (op->xltype == xltypeInt) {
            return protocol::CreateAny(builder, protocol::AnyValue::Int,
                                       protocol::CreateInt(builder, op->val.w).Union());
        } else if (op->xltype == xltypeBool) {
            return protocol::CreateAny(builder, protocol::AnyValue::Bool,
                                       protocol::CreateBool(builder, op->val.xbool).Union());
        } else if (op->xltype == xltypeStr) {
            return protocol::CreateAny(builder, protocol::AnyValue::Str,
                                       protocol::CreateStr(builder, builder.CreateString(ConvertExcelString(op->val.str)))
                                           .Union());
        } else if (op->xltype == xltypeErr) {
            return protocol::CreateAny(builder, protocol::AnyValue::Err,
                                       protocol::CreateErr(builder, ExcelErrorToProtocol(op->val.err)).Union());
        } else if (op->xltype & (xltypeRef | xltypeSRef)) {
            return protocol::CreateAny(builder, protocol::AnyValue::Range, ConvertRange(op, builder).Union());

        } else if (op->xltype & xltypeMulti) {
            return ConvertMultiToAny(*op, builder);
        } else if (op->xltype & (xltypeMissing | xltypeNil)) {
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
                 size_t count = (size_t)rows * cols;

                 if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max() ||
                     !ng->data() || ng->data()->size() < count) {
                     LPXLOPER12 op = NewXLOPER12();
                     op->xltype = xltypeErr | xlbitDLLFree;
                     op->val.err = xlerrValue;
                     return op;
                 }

                 // Check for allocation overflow (size_t)
                 if (count > SIZE_MAX / sizeof(XLOPER12)) {
                     LPXLOPER12 op = NewXLOPER12();
                     op->xltype = xltypeErr | xlbitDLLFree;
                     op->val.err = xlerrValue;
                     return op;
                 }

                 LPXLOPER12 op = NewXLOPER12();
                 op->xltype = xltypeMulti | xlbitDLLFree;
                 op->val.array.rows = rows;
                 op->val.array.columns = cols;

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
                LPXLOPER12 op = NewXLOPER12();
                op->xltype = xltypeErr | xlbitDLLFree;
                op->val.err = xlerrNA;
                return op;
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
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }
}

LPXLOPER12 RangeToXLOPER12(const protocol::Range* range) {
    if (!range) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    // Validate refs existence and count limits (XLOPER12 ref count is WORD/16-bit)
    if (!range->refs() || range->refs()->size() > 65535) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    try {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeRef | xlbitDLLFree;

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
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }
}

LPXLOPER12 GridToXLOPER12(const protocol::Grid* grid) {
    if (!grid) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    int rows = grid->rows();
    int cols = grid->cols();

    if (rows < 0 || cols < 0) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    if (cols > 0 && (size_t)rows > SIZE_MAX / (size_t)cols) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    size_t count = (size_t)rows * (size_t)cols;

    if (count > (size_t)std::numeric_limits<int>::max() ||
        !grid->data() || grid->data()->size() != count) {
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    // Check for allocation overflow (size_t)
    if (count > SIZE_MAX / sizeof(XLOPER12)) {
         LPXLOPER12 op = NewXLOPER12();
         op->xltype = xltypeErr | xlbitDLLFree;
         op->val.err = xlerrValue;
         return op;
    }

    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeMulti | xlbitDLLFree;
    op->val.array.rows = rows;
    op->val.array.columns = cols;

    // RAII guard to clean up if exception happens or return early
    ScopeGuard guard([&]() {
        if (op->val.array.lparray) {
            for(size_t i=0; i<count; ++i) {
                 if (op->val.array.lparray[i].xltype == xltypeStr && op->val.array.lparray[i].val.str) {
                     delete[] op->val.array.lparray[i].val.str;
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
                    const auto* fbStr = scalar->val_as_Str()->val();
                    if (!fbStr) {
                        cell.val.str = new XCHAR[2];
                        cell.val.str[0] = 0;
                        cell.val.str[1] = 0;
                    } else {
                        const char* utf8 = fbStr->c_str();
                        size_t realLen = fbStr->size();

                        // Optimization/Security: Excel 12 strings are limited to 32767 chars.
                        // If the input is huge, we don't need to convert all of it just to truncate it.
                        // 32767 chars can be at most ~132KB (4 bytes per char) in UTF-8.
                        // We clamp the input to a safe margin (200,000 bytes) to prevent allocating
                        // huge temporary buffers (e.g. 20MB) for strings that will be truncated anyway.
                        // This also prevents integer overflow when casting size_t to int.
                        if (realLen > 200000) {
                            realLen = 200000;
                        }
                        int utf8Len = static_cast<int>(realLen);

                        // CP_UTF8 = 65001
                        // Optimization: Try to convert using a stack buffer first to avoid double API call.
                        // Most Excel strings are small.
                        XCHAR stackBuf[256];
                        int needed = 0;

                        if (utf8Len < 256) {
                            needed = MultiByteToWideChar(65001, 0, utf8, utf8Len, stackBuf, 256);
                        }

                        if (needed > 0) {
                            // Successful conversion to stack buffer
                            cell.val.str = new XCHAR[needed + 2];
                            std::memcpy(cell.val.str + 1, stackBuf, needed * sizeof(XCHAR));
                            cell.val.str[0] = (XCHAR)needed;
                            cell.val.str[needed + 1] = 0;
                        } else {
                            // Fallback to double-call (or string too long for stack buffer)
                            needed = MultiByteToWideChar(65001, 0, utf8, utf8Len, NULL, 0);

                            // Safety check: Don't allocate huge memory for strings.
                            if (needed < 0 || needed > 10000000 || (size_t)needed > SIZE_MAX / sizeof(XCHAR) - 2) {
                                cell.val.str = new XCHAR[2];
                                cell.val.str[0] = 0;
                                cell.val.str[1] = 0;
                            } else {
                                if (needed > 32767) {
                                    XCHAR* temp = new XCHAR[needed + 2];
                                    MultiByteToWideChar(65001, 0, utf8, utf8Len, temp + 1, needed);

                                    cell.val.str = new XCHAR[32767 + 2];
                                    std::memcpy(cell.val.str + 1, temp + 1, 32767 * sizeof(XCHAR));

                                    delete[] temp;

                                    cell.val.str[0] = 32767;
                                    cell.val.str[32767 + 1] = 0;
                                } else {
                                    cell.val.str = new XCHAR[needed + 2];
                                    if (needed > 0) {
                                        MultiByteToWideChar(65001, 0, utf8, utf8Len, cell.val.str + 1, needed);
                                    }
                                    cell.val.str[0] = (XCHAR)needed;
                                    cell.val.str[needed + 1] = 0;
                                }
                            }
                        }
                    }
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
        LPXLOPER12 op = NewXLOPER12();
        op->xltype = xltypeErr | xlbitDLLFree;
        op->val.err = xlerrValue;
        return op;
    }

    guard.Dismiss();
    return op;
}

FP12* NumGridToFP12(const protocol::NumGrid* grid) {
    try {
        if (!grid) return NewFP12(0, 0);
        int rows = grid->rows();
        int cols = grid->cols();

        if (rows < 0 || cols < 0) {
            return NewFP12(0, 0);
        }

        if (cols > 0 && (size_t)rows > SIZE_MAX / (size_t)cols) {
            return NewFP12(0, 0);
        }

        size_t count = (size_t)rows * (size_t)cols;

        if (count > (size_t)std::numeric_limits<int>::max() ||
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
