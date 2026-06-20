// R33 regression suite: the single Excel Pascal wide-string writer
// (WritePascalWString) and the five call sites that now delegate to it.
//
// Boundary matrix exercised per encoder:
//   len 0  /  len < 32767  /  len == 32767  /  len > 32767 (clamp)
// plus: dst[0] length value, body fidelity, trailing-NUL position,
//       embedded NUL preservation, non-BMP surrogate-pair fidelity,
//       and over-read / over-write guards (canary words around buffers).
//
// No flatbuffers dependency: this exercises pascalstr.h + the string paths of
// ScopedXLOPER12, NewExcelString, and Utf8ToExcelString directly.

#include <cassert>
#include <cstring>
#include <iostream>
#include <string>
#include <vector>

#include <windows.h>
HINSTANCE g_hModule = NULL;

#include "types/pascalstr.h"
#include "types/ScopedXLOPER12.h"
#include "types/mem.h"
#include "types/utility.h"

extern "C" void __stdcall xlAutoFree12(LPXLOPER12 p);

static int g_failures = 0;
#define CHECK(cond)                                                            \
    do {                                                                       \
        if (!(cond)) {                                                         \
            std::cerr << "FAIL: " << #cond << " @ " << __FILE__ << ":"         \
                      << __LINE__ << std::endl;                                \
            ++g_failures;                                                      \
        }                                                                      \
    } while (0)

// ---------------------------------------------------------------------------
// 1. WritePascalWString core boundary behavior, with canary guards.
// ---------------------------------------------------------------------------

// Allocate a buffer with one canary word before and after the region the
// writer is allowed to touch, so an off-by-one over-read/write is caught.
static void run_writer_case(size_t srcLen) {
    const XCHAR kCanary = (XCHAR)0xBEEF;
    size_t bodyLen = WritePascalWBufferLen(srcLen); // prefix+body+NUL
    size_t expectedClamp = srcLen > kMaxExcelStringLen ? kMaxExcelStringLen : srcLen;

    // Source body filled with a recognizable, position-dependent pattern.
    std::vector<XCHAR> src(srcLen ? srcLen : 1);
    for (size_t i = 0; i < srcLen; ++i) {
        src[i] = (XCHAR)(L'A' + (i % 26));
    }

    // [canary][ writer region (bodyLen) ][canary]
    std::vector<XCHAR> buf(bodyLen + 2, kCanary);
    XCHAR* dst = buf.data() + 1;

    size_t written = WritePascalWString(dst, srcLen ? src.data() : nullptr, srcLen);

    CHECK(written == expectedClamp);
    CHECK((size_t)dst[0] == expectedClamp);          // length prefix
    for (size_t i = 0; i < expectedClamp; ++i) {     // body fidelity
        CHECK(dst[1 + i] == (XCHAR)(L'A' + (i % 26)));
    }
    CHECK(dst[expectedClamp + 1] == 0);              // trailing NUL position
    // Canaries intact -> no under/over-run.
    CHECK(buf.front() == kCanary);
    CHECK(buf.back() == kCanary);
}

static void TestWriterBoundaries() {
    run_writer_case(0);
    run_writer_case(1);
    run_writer_case(5);
    run_writer_case(255);
    run_writer_case(256);
    run_writer_case(kMaxExcelStringLen - 1); // 32766
    run_writer_case(kMaxExcelStringLen);     // 32767 exact -> kept
    run_writer_case(kMaxExcelStringLen + 1); // 32768 -> clamp to 32767
    run_writer_case(40000);                  // clamp to 32767

    // Embedded NUL preserved (copy is length-driven, not NUL-driven).
    {
        XCHAR src[5] = {L'A', 0, L'B', 0, L'C'};
        XCHAR dst[8];
        size_t w = WritePascalWString(dst, src, 5);
        CHECK(w == 5);
        CHECK((size_t)dst[0] == 5);
        CHECK(dst[1] == L'A');
        CHECK(dst[2] == 0);
        CHECK(dst[3] == L'B');
        CHECK(dst[4] == 0);
        CHECK(dst[5] == L'C');
        CHECK(dst[6] == 0); // trailing NUL
    }

    // Non-BMP (surrogate pair) input: U+1F600 = 0xD83D 0xDE00, counted as 2
    // UTF-16 code units, copied verbatim.
    {
        XCHAR src[2] = {(XCHAR)0xD83D, (XCHAR)0xDE00};
        XCHAR dst[5];
        size_t w = WritePascalWString(dst, src, 2);
        CHECK(w == 2);
        CHECK((size_t)dst[0] == 2);
        CHECK(dst[1] == (XCHAR)0xD83D);
        CHECK(dst[2] == (XCHAR)0xDE00);
        CHECK(dst[3] == 0);
    }

    std::cout << "TestWriterBoundaries done" << std::endl;
}

// ---------------------------------------------------------------------------
// 2. ScopedXLOPER12::SetString (via wstring ctor) boundaries.
// ---------------------------------------------------------------------------
static void check_scoped_from_wstring(size_t len) {
    size_t expected = len > kMaxExcelStringLen ? kMaxExcelStringLen : len;
    std::wstring s(len, L'x');
    ScopedXLOPER12 op(s);
    LPXLOPER12 p = op;
    if (len == 0) {
        // Empty wstring still yields a 0-length xltypeStr (not Missing): wstring
        // ctor passes c_str() which is non-null.
        CHECK(p->xltype == xltypeStr);
        CHECK((size_t)p->val.str[0] == 0);
        CHECK(p->val.str[1] == 0);
        return;
    }
    CHECK(p->xltype == xltypeStr);
    CHECK((size_t)p->val.str[0] == expected);
    CHECK(p->val.str[1] == L'x');
    CHECK(p->val.str[expected] == L'x');
    CHECK(p->val.str[expected + 1] == 0); // NUL terminated
}

static void TestScopedSetString() {
    check_scoped_from_wstring(0);
    check_scoped_from_wstring(5);
    check_scoped_from_wstring(kMaxExcelStringLen - 1);
    check_scoped_from_wstring(kMaxExcelStringLen);
    check_scoped_from_wstring(kMaxExcelStringLen + 1);
    check_scoped_from_wstring(33000);

    // null wchar_t* -> Missing (unchanged behavior)
    {
        ScopedXLOPER12 op((const wchar_t*)nullptr);
        LPXLOPER12 p = op;
        CHECK(p->xltype == xltypeMissing);
    }
    std::cout << "TestScopedSetString done" << std::endl;
}

// ---------------------------------------------------------------------------
// 3. ScopedXLOPER12(const XLOPER12*) string-copy ctor, incl. normalization of
//    a MALFORMED oversized length prefix (intentional behavior change vs the
//    old open-coded path, which copied the bogus prefix verbatim).
// ---------------------------------------------------------------------------
static void TestScopedFromXloper() {
    // Well-formed source.
    {
        XLOPER12 src;
        src.xltype = xltypeStr;
        wchar_t buf[7] = {5, L'H', L'e', L'l', L'l', L'o', 0};
        src.val.str = buf;
        ScopedXLOPER12 op(&src);
        LPXLOPER12 p = op;
        CHECK(p->xltype == xltypeStr);
        CHECK((size_t)p->val.str[0] == 5);
        CHECK(std::wcsncmp(p->val.str + 1, L"Hello", 5) == 0);
        CHECK(p->val.str[6] == 0);
    }
    // Exactly 32767.
    {
        std::vector<wchar_t> buf(kMaxExcelStringLen + 2, L'z');
        buf[0] = (wchar_t)kMaxExcelStringLen;
        buf[kMaxExcelStringLen + 1] = 0;
        XLOPER12 src;
        src.xltype = xltypeStr;
        src.val.str = buf.data();
        ScopedXLOPER12 op(&src);
        LPXLOPER12 p = op;
        CHECK((size_t)p->val.str[0] == kMaxExcelStringLen);
        CHECK(p->val.str[kMaxExcelStringLen] == L'z');
        CHECK(p->val.str[kMaxExcelStringLen + 1] == 0);
    }
    // MALFORMED: prefix claims 40000 but allocate a big enough body so the
    // clamped read stays in-bounds. NEW behavior: prefix is normalized to
    // 32767 (old code left dst[0]==40000, a latent off-by-one/overrun bug for
    // any downstream reader trusting the prefix).
    {
        std::vector<wchar_t> buf(40000 + 2, L'q');
        buf[0] = (wchar_t)40000; // bogus oversized length
        XLOPER12 src;
        src.xltype = xltypeStr;
        src.val.str = buf.data();
        ScopedXLOPER12 op(&src);
        LPXLOPER12 p = op;
        CHECK((size_t)p->val.str[0] == kMaxExcelStringLen); // normalized
        CHECK(p->val.str[1] == L'q');
        CHECK(p->val.str[kMaxExcelStringLen] == L'q');
        CHECK(p->val.str[kMaxExcelStringLen + 1] == 0);
    }
    // null payload -> Missing (unchanged).
    {
        XLOPER12 src;
        src.xltype = xltypeStr;
        src.val.str = nullptr;
        ScopedXLOPER12 op(&src);
        LPXLOPER12 p = op;
        CHECK(p->xltype == xltypeMissing);
    }
    std::cout << "TestScopedFromXloper done" << std::endl;
}

// ---------------------------------------------------------------------------
// 4. NewExcelString boundaries (DLL-owned buffer; freed via xlAutoFree12).
// ---------------------------------------------------------------------------
static void check_new_excel_string(size_t len) {
    size_t expected = len > kMaxExcelStringLen ? kMaxExcelStringLen : len;
    std::wstring s(len, L'k');
    LPXLOPER12 p = NewExcelString(s);
    CHECK(p != nullptr);
    CHECK(p->xltype == (xltypeStr | xlbitDLLFree));
    CHECK((size_t)p->val.str[0] == expected);
    if (expected > 0) {
        CHECK(p->val.str[1] == L'k');
        CHECK(p->val.str[expected] == L'k');
    }
    CHECK(p->val.str[expected + 1] == 0); // NUL terminated
    xlAutoFree12(p);                       // exercises FreeDllOwnedContents path
}

static void TestNewExcelString() {
    check_new_excel_string(0);
    check_new_excel_string(5);
    check_new_excel_string(kMaxExcelStringLen - 1);
    check_new_excel_string(kMaxExcelStringLen);
    check_new_excel_string(kMaxExcelStringLen + 1);
    check_new_excel_string(50000);
    std::cout << "TestNewExcelString done" << std::endl;
}

// ---------------------------------------------------------------------------
// 5. Utf8ToExcelString boundaries (caller-owned buffer; ASCII so byte==char).
//    Covers all three internal paths: stack (<256), heap in-range, >32767.
// ---------------------------------------------------------------------------
static void check_utf8(size_t len) {
    // ASCII 'M' => one UTF-16 unit per byte, so post-conversion length == len
    // before clamping, EXCEPT that the function clamps the *input bytes* to
    // 200000 first. For len <= 200000 the char count is exactly len.
    size_t expected = len > kMaxExcelStringLen ? kMaxExcelStringLen : len;
    std::string s(len, 'M');
    XCHAR* out = nullptr;
    Utf8ToExcelString(s.c_str(), out);
    CHECK(out != nullptr);
    CHECK((size_t)out[0] == expected);
    if (expected > 0) {
        CHECK(out[1] == L'M');
        CHECK(out[expected] == L'M');
    }
    CHECK(out[expected + 1] == 0); // NUL terminated
    delete[] out;                  // caller owns
}

static void TestUtf8ToExcelString() {
    check_utf8(0);
    check_utf8(5);
    check_utf8(255);                       // stack-path boundary
    check_utf8(256);                       // first heap path
    check_utf8(kMaxExcelStringLen - 1);
    check_utf8(kMaxExcelStringLen);        // exact
    check_utf8(kMaxExcelStringLen + 1);    // clamp
    check_utf8(60000);                     // clamp

    // Non-BMP: U+1F600 in UTF-8 (F0 9F 98 80) -> 2 UTF-16 units.
    {
        const char* emoji = "\xF0\x9F\x98\x80";
        XCHAR* out = nullptr;
        Utf8ToExcelString(emoji, out);
        CHECK(out != nullptr);
        CHECK((size_t)out[0] == 2);
        CHECK(out[1] == (XCHAR)0xD83D);
        CHECK(out[2] == (XCHAR)0xDE00);
        CHECK(out[3] == 0);
        delete[] out;
    }

    // null input -> empty string (unchanged behavior).
    {
        XCHAR* out = nullptr;
        Utf8ToExcelString(nullptr, out);
        CHECK(out != nullptr);
        CHECK(out[0] == 0);
        delete[] out;
    }
    std::cout << "TestUtf8ToExcelString done" << std::endl;
}

int main() {
    TestWriterBoundaries();
    TestScopedSetString();
    TestScopedFromXloper();
    TestNewExcelString();
    TestUtf8ToExcelString();

    if (g_failures) {
        std::cerr << g_failures << " CHECK(s) FAILED" << std::endl;
        return 1;
    }
    std::cout << "All pascalstr regression tests passed" << std::endl;
    return 0;
}
