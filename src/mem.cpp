#include "types/mem.h"
#include "types/ObjectPool.h"
#include "types/ScopeGuard.h"
#include <mutex>
#include <vector>
#include <cstring> // For memset, memcpy

static ObjectPool<XLOPER12> xloperPool;

LPXLOPER12 NewXLOPER12() {
    LPXLOPER12 p = xloperPool.Acquire();
    // Initialize to zero.
    std::memset(p, 0, sizeof(XLOPER12));
    return p;
}

void ReleaseXLOPER12(LPXLOPER12 p) {
    if (p) {
        xloperPool.Release(p);
    }
}

LPXLOPER12 NewExcelString(const std::wstring& str) {
    LPXLOPER12 p = NewXLOPER12();
    p->xltype = xltypeStr | xlbitDLLFree;

    ScopeGuard guard([&]() {
        ReleaseXLOPER12(p);
    });

    size_t len = str.length();
    if (len > 32767) len = 32767;

    wchar_t* buffer = new wchar_t[len + 2];
    buffer[0] = (wchar_t)len;
    if (len > 0) std::memcpy(buffer + 1, str.data(), len * sizeof(wchar_t));
    buffer[len + 1] = 0; // Null terminate for safety, though not strictly required for Pascal string

    p->val.str = buffer;
    guard.Dismiss();
    return p;
}

// Thread-local storage for FP12 return values
struct FP12Buffer {
    std::vector<char> data; // Stores FP12 header + doubles
};

// Keep a few buffers per thread to allow safe returns without immediate overwrites.
//
// LIMITATION: this is a fixed-size per-thread ring of kRingBufferSize (8)
// buffers. A pointer returned by NewFP12 stays valid only until the same thread
// calls NewFP12 kRingBufferSize more times — the 9th subsequent call wraps back
// to the same FP12Buffer and may resize-reallocate it, invalidating the older
// pointer. This is safe for the intended usage (Excel copies each returned FP12
// before the wrapper returns, so only a couple are live at once), but a caller
// that holds >8 NewFP12 results simultaneously (e.g. deeply nested grid
// building on one thread) will silently read overwritten/freed data. If that
// pattern ever appears, switch the call site to a caller-owned allocation
// rather than enlarging this ring.
static const int kRingBufferSize = 8;
thread_local int fpRingIndex = 0;
thread_local FP12Buffer fpRingBuffers[kRingBufferSize];

FP12* NewFP12(int rows, int cols) {
    FP12Buffer& buf = fpRingBuffers[fpRingIndex];
    fpRingIndex = (fpRingIndex + 1) % kRingBufferSize;

    // Calculate required size: 2 ints + rows*cols doubles
    size_t size = sizeof(INT32) * 2 + (size_t)rows * cols * sizeof(double);
    if (buf.data.size() < size) {
        buf.data.resize(size);
    }

    FP12* fp = reinterpret_cast<FP12*>(buf.data.data());
    fp->rows = rows;
    fp->columns = cols;
    return fp;
}

TYPES_EXCEL_CALLBACK xlAutoFree12(LPXLOPER12 p) {
    if (!p) return;

    // Check if the XLOPER12 itself is marked for DLL freeing
    // (Usually this function is only called if xlbitDLLFree is set on p->xltype)

    if (p->xltype & xltypeStr) {
        if (p->val.str) {
            delete[] p->val.str;
            p->val.str = nullptr;
        }
    }
    else if (p->xltype & xltypeMulti) {
         if (p->val.array.lparray) {
             size_t count = (size_t)p->val.array.rows * p->val.array.columns;
             // Ownership contract: an element string is only delete[]'d when
             // the element carries xlbitDLLFree — the explicit marker set by
             // GridToXLOPER12 (src/converters.cpp) on strings it allocates
             // via Utf8ToExcelString. An element WITHOUT the bit (e.g. an
             // Excel-owned or aliased pointer placed into a cell by other
             // code) is left alone. Excel ignores xlbitDLLFree on inner
             // elements, so the bit is purely our ownership marker; our own
             // readers (ConvertScalar/ConvertMultiToAny/ConvertAny) mask it
             // before type dispatch.
             for(size_t i=0; i<count; ++i) {
                 LPXLOPER12 elem = &p->val.array.lparray[i];
                 if ((elem->xltype & xltypeStr) && (elem->xltype & xlbitDLLFree) && elem->val.str) {
                      delete[] elem->val.str;
                      elem->val.str = nullptr;
                 }
             }
             delete[] p->val.array.lparray;
         }
    }
    else if (p->xltype & xltypeRef) {
        if (p->val.mref.lpmref) {
            delete[] (char*)p->val.mref.lpmref;
            p->val.mref.lpmref = nullptr;
        }
    }

    // Finally, release the XLOPER12 struct itself back to the pool
    xloperPool.Release(p);
}
