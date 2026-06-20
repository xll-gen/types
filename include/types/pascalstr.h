#pragma once

#include <windows.h>

#include "types/xlcall.h"
#include <cstddef>
#include <cstring>

// =============================================================================
// Single writer for the Excel "Pascal" wide-string layout (XLOPER12 xltypeStr).
// =============================================================================
//
// Excel 12 string layout (`XCHAR*`, i.e. `wchar_t*` on Windows):
//   buf[0]            = length, stored as one XCHAR  (max 32767)
//   buf[1 .. len]     = the body characters
//   buf[len+1]        = optional terminating NUL (NOT required by Excel, but
//                       every producer in this repo writes it for safety)
//
// `kMaxExcelStringLen` is the hard Excel limit. A length of exactly 32767 is
// LEGAL and is copied through verbatim; only 32768+ is clamped.
//
// `WritePascalWString` is the ONE place that encodes this layout. Every call
// site is responsible for allocating a buffer of AT LEAST
// `WritePascalWBufferLen(srcLen)` XCHARs and keeps its own ownership /
// free strategy — this helper never allocates and never frees.
//
// ABI: this only formats bytes into a caller-provided buffer; the XLOPER12
// struct layout is untouched.

inline constexpr size_t kMaxExcelStringLen = 32767;

// Single clamp used by BOTH the buffer-sizing and the writer below, so the two
// can never desync (a divergence would size a buffer for one length and write
// another — a heap overflow). Do not inline the ternary at the call sites.
inline constexpr size_t ClampExcelStringLen(size_t len) {
    return len > kMaxExcelStringLen ? kMaxExcelStringLen : len;
}

// Number of XCHARs a buffer must hold to store a Pascal string whose source
// body is `srcLen` characters long: clamp(srcLen) length-prefix + body + NUL.
inline size_t WritePascalWBufferLen(size_t srcLen) {
    return ClampExcelStringLen(srcLen) + 2; // [0]=len prefix, body, trailing NUL
}

// Encode `src[0 .. len)` into the Excel Pascal wide-string `dst`.
//
//   - clamps `len` to kMaxExcelStringLen (32767); a `len` of exactly 32767 is
//     kept, only > 32767 is truncated,
//   - writes dst[0] = (XCHAR)clampedLen,
//   - copies the clamped body to dst[1 .. clampedLen],
//   - NUL-terminates at dst[clampedLen + 1].
//
// `dst` must point to at least `WritePascalWBufferLen(len)` XCHARs. `src` may be
// null only when `len` is 0. Returns the clamped length actually written (handy
// for callers that need to know the truncated count).
//
// Embedded NULs inside the body are preserved (the copy is length-driven, not
// NUL-driven). Surrogate pairs (non-BMP input) are copied as their constituent
// UTF-16 code units — `len` is a code-unit count, exactly as Excel expects.
inline size_t WritePascalWString(XCHAR* dst, const XCHAR* src, size_t len) {
    size_t clampedLen = ClampExcelStringLen(len);
    dst[0] = (XCHAR)clampedLen;
    if (clampedLen > 0) {
        std::memcpy(dst + 1, src, clampedLen * sizeof(XCHAR));
    }
    dst[clampedLen + 1] = 0; // trailing NUL (safety; not required by Excel)
    return clampedLen;
}
