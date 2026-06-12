// test_roundtrip.cpp
//
// XLOPER12 <-> FlatBuffer round-trip fixture for every converter exposed by
// include/types/converters.h. Mirrors the Go-side coverage in
// go/protocol/deepcopy_test.go: each test builds an input, runs it through
// Convert*() -> Finish() -> GetRoot() -> *ToXLOPER12() and asserts the
// result is byte-equal (modulo the xlbitDLLFree flag the C++ side adds on
// return).
//
// Tests are organized one-per-converter:
//   - Scalar:  Int / Num / Bool / Str / Err (+ NaN, Inf, INT_MIN, UTF-16
//              non-ASCII content)
//   - Grid:    2x2 mixed types, 1x1, empty
//   - NumGrid: 2x2 doubles, NaN/Inf preserved bit-for-bit
//   - Range:   single rectangular ref
//   - Any:     each variant routed through ConvertAny()/AnyToXLOPER12()
//
// Closes the last open item in AGENTS.md "Known Improvement Backlog"
// (v0.2.6).

#include <cassert>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iostream>
#include <limits>
#include <sstream>
#include <string>
#include <vector>

#ifdef _WIN32
#include <windows.h>
// Some Excel symbols want a module handle even in tests; the existing
// test_converters.cpp uses the same idiom.
HINSTANCE g_hModule = NULL;
#endif

#include "types/converters.h"
#include "types/mem.h"
#include "types/utility.h"

extern "C" void __stdcall xlAutoFree12(LPXLOPER12 p);

// ----------------------------------------------------------------------------
// Tiny test harness. We want each named test to print pass/fail + line of the
// failing assertion, and to keep going on failure so CI shows the full damage
// rather than the first miss.

static int g_failures = 0;
static const char* g_current_test = "";

#define RT_CHECK(cond)                                                     \
    do {                                                                   \
        if (!(cond)) {                                                     \
            ++g_failures;                                                  \
            std::cerr << "  FAIL [" << g_current_test << "] " << __FILE__  \
                      << ":" << __LINE__ << "  " << #cond << std::endl;    \
        }                                                                  \
    } while (0)

#define RT_CHECK_EQ(a, b)                                                  \
    do {                                                                   \
        auto _av = (a);                                                    \
        auto _bv = (b);                                                    \
        if (!(_av == _bv)) {                                               \
            ++g_failures;                                                  \
            std::cerr << "  FAIL [" << g_current_test << "] " << __FILE__  \
                      << ":" << __LINE__ << "  " << #a << " != " << #b     \
                      << "  (got " << _av << " vs " << _bv << ")"          \
                      << std::endl;                                        \
        }                                                                  \
    } while (0)

#define RT_RUN(fn)                                                         \
    do {                                                                   \
        g_current_test = #fn;                                              \
        int before = g_failures;                                           \
        fn();                                                              \
        if (g_failures == before) {                                        \
            std::cout << "  PASS " << #fn << std::endl;                    \
        }                                                                  \
    } while (0)

// ----------------------------------------------------------------------------
// Helpers.

// Build a Pascal-style XCHAR string suitable for assigning to
// XLOPER12::val.str. Caller owns the returned buffer (delete[]).
static XCHAR* MakePascalString(const std::wstring& s) {
    size_t len = s.size();
    if (len > 32767) len = 32767;
    XCHAR* buf = new XCHAR[len + 2];
    buf[0] = static_cast<XCHAR>(len);
    for (size_t i = 0; i < len; ++i) buf[i + 1] = s[i];
    buf[len + 1] = 0;
    return buf;
}

// Compare two doubles for bitwise equality. Plain == is wrong for NaN.
static bool BitEqualDouble(double a, double b) {
    uint64_t ua, ub;
    std::memcpy(&ua, &a, sizeof(ua));
    std::memcpy(&ub, &b, sizeof(ub));
    return ua == ub;
}

// xltype of a result XLOPER12 with the xlbitDLLFree flag stripped — that
// flag is always set by the C++ side on returned heap objects and is not
// part of the logical type identity.
static DWORD CoreType(const XLOPER12& op) {
    return op.xltype & ~static_cast<DWORD>(xlbitDLLFree);
}

// Round-trip a single XLOPER12 through Any. Caller must xlAutoFree12() the
// result.
static LPXLOPER12 RoundTripAny(LPXLOPER12 src) {
    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(src, builder);
    builder.Finish(off);
    auto* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    return AnyToXLOPER12(root);
}

// ----------------------------------------------------------------------------
// Scalar round-trips.

static void TestScalar_Int_Basic() {
    XLOPER12 in{};
    in.xltype = xltypeInt;
    in.val.w = 42;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.w, 42);
    xlAutoFree12(out);
}

static void TestScalar_Int_Negative() {
    XLOPER12 in{};
    in.xltype = xltypeInt;
    in.val.w = -1;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.w, -1);
    xlAutoFree12(out);
}

static void TestScalar_Int_Min() {
    XLOPER12 in{};
    in.xltype = xltypeInt;
    in.val.w = std::numeric_limits<int>::min();
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.w, std::numeric_limits<int>::min());
    xlAutoFree12(out);
}

static void TestScalar_Int_Max() {
    XLOPER12 in{};
    in.xltype = xltypeInt;
    in.val.w = std::numeric_limits<int>::max();
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.w, std::numeric_limits<int>::max());
    xlAutoFree12(out);
}

static void TestScalar_Num_Basic() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = 3.14159265358979;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.num, 3.14159265358979));
    xlAutoFree12(out);
}

static void TestScalar_Num_Zero() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = 0.0;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.num, 0.0));
    xlAutoFree12(out);
}

static void TestScalar_Num_NegativeZero() {
    // Documented protocol limitation: `table Num { val: double; }` has a
    // default of 0.0. FlatBuffers' default-value elision treats -0.0 as
    // equal to 0.0 and omits the field, so on read -0.0 round-trips to
    // +0.0. The vector-of-double path (NumGrid) is unaffected and *does*
    // preserve the sign bit — see TestNumGrid_WithSpecials below.
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = -0.0;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    // Value is numerically zero (sign lost in scalar table encoding).
    RT_CHECK_EQ(out->val.num, 0.0);
    xlAutoFree12(out);
}

static void TestScalar_Num_NaN() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = std::numeric_limits<double>::quiet_NaN();
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(std::isnan(out->val.num));
    xlAutoFree12(out);
}

static void TestScalar_Num_PosInf() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = std::numeric_limits<double>::infinity();
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(std::isinf(out->val.num) && out->val.num > 0);
    xlAutoFree12(out);
}

static void TestScalar_Num_NegInf() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = -std::numeric_limits<double>::infinity();
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(std::isinf(out->val.num) && out->val.num < 0);
    xlAutoFree12(out);
}

static void TestScalar_Str_Ascii() {
    XLOPER12 in{};
    in.xltype = xltypeStr;
    std::wstring s = L"Hello, XLL!";
    in.val.str = MakePascalString(s);

    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeStr));
    RT_CHECK(out->val.str != nullptr);
    if (out->val.str) {
        size_t outLen = static_cast<size_t>(out->val.str[0]);
        RT_CHECK_EQ(outLen, s.size());
        bool match = (outLen == s.size());
        for (size_t i = 0; match && i < outLen; ++i) {
            if (out->val.str[i + 1] != s[i]) match = false;
        }
        RT_CHECK(match);
    }
    delete[] in.val.str;
    xlAutoFree12(out);
}

static void TestScalar_Str_Empty() {
    XLOPER12 in{};
    in.xltype = xltypeStr;
    in.val.str = MakePascalString(L"");

    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeStr));
    RT_CHECK(out->val.str != nullptr);
    if (out->val.str) {
        RT_CHECK_EQ(static_cast<int>(out->val.str[0]), 0);
    }
    delete[] in.val.str;
    xlAutoFree12(out);
}

static void TestScalar_Str_Unicode() {
    XLOPER12 in{};
    in.xltype = xltypeStr;
    // Korean + Greek + CJK to exercise multi-byte UTF-8 round-trip through
    // the FlatBuffers UTF-8 string field.
    std::wstring s = L"한글 αβγ 中文";
    in.val.str = MakePascalString(s);

    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeStr));
    RT_CHECK(out->val.str != nullptr);
    if (out->val.str) {
        size_t outLen = static_cast<size_t>(out->val.str[0]);
        RT_CHECK_EQ(outLen, s.size());
        bool match = (outLen == s.size());
        for (size_t i = 0; match && i < outLen; ++i) {
            if (out->val.str[i + 1] != s[i]) match = false;
        }
        RT_CHECK(match);
    }
    delete[] in.val.str;
    xlAutoFree12(out);
}

static void TestScalar_Bool_True() {
    XLOPER12 in{};
    in.xltype = xltypeBool;
    in.val.xbool = 1;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeBool));
    RT_CHECK(out->val.xbool != 0);
    xlAutoFree12(out);
}

static void TestScalar_Bool_False() {
    XLOPER12 in{};
    in.xltype = xltypeBool;
    in.val.xbool = 0;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeBool));
    RT_CHECK_EQ(out->val.xbool, 0);
    xlAutoFree12(out);
}

static void TestScalar_Err_Value() {
    XLOPER12 in{};
    in.xltype = xltypeErr;
    in.val.err = xlerrValue;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeErr));
    RT_CHECK_EQ(out->val.err, xlerrValue);
    xlAutoFree12(out);
}

static void TestScalar_Err_Num() {
    XLOPER12 in{};
    in.xltype = xltypeErr;
    in.val.err = xlerrNum;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeErr));
    RT_CHECK_EQ(out->val.err, xlerrNum);
    xlAutoFree12(out);
}

static void TestScalar_Err_Div0() {
    XLOPER12 in{};
    in.xltype = xltypeErr;
    in.val.err = xlerrDiv0;
    LPXLOPER12 out = RoundTripAny(&in);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeErr));
    RT_CHECK_EQ(out->val.err, xlerrDiv0);
    xlAutoFree12(out);
}

// Direct ConvertScalar exercise (not via Any) — checks the Scalar union
// path independently. Goes through Grid 1x1 since there's no public
// ScalarToXLOPER12 entry point: we wrap the Scalar in a Grid for the
// reverse leg, which is what AnyValue::Grid does inside AnyToXLOPER12.
static void TestConvertScalar_Direct_Num() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 cell{};
    cell.xltype = xltypeNum;
    cell.val.num = 7.5;

    auto scalarOff = ConvertScalar(cell, builder);
    std::vector<flatbuffers::Offset<protocol::Scalar>> elems = {scalarOff};
    auto vec = builder.CreateVector(elems);
    auto gridOff = protocol::CreateGrid(builder, 1, 1, vec);
    builder.Finish(gridOff);

    auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 out = GridToXLOPER12(grid);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 1);
    RT_CHECK_EQ(out->val.array.columns, 1);
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[0].val.num, 7.5));
    xlAutoFree12(out);
}

// ----------------------------------------------------------------------------
// Grid round-trips. ConvertGrid takes a Multi XLOPER12 and produces a
// protocol::Grid; GridToXLOPER12 reverses it.

static void TestGrid_2x2_Mixed() {
    // 2x2 grid: [ [int 1, num 2.5], [bool true, str "ok"] ]
    XLOPER12 cells[4];
    cells[0].xltype = xltypeInt;
    cells[0].val.w = 1;
    cells[1].xltype = xltypeNum;
    cells[1].val.num = 2.5;
    cells[2].xltype = xltypeBool;
    cells[2].val.xbool = 1;
    cells[3].xltype = xltypeStr;
    cells[3].val.str = MakePascalString(L"ok");

    XLOPER12 multi{};
    multi.xltype = xltypeMulti;
    multi.val.array.rows = 2;
    multi.val.array.columns = 2;
    multi.val.array.lparray = cells;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertGrid(&multi, builder);
    builder.Finish(off);
    auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 out = GridToXLOPER12(grid);

    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 2);
    RT_CHECK_EQ(out->val.array.columns, 2);
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.array.lparray[0].val.w, 1);
    RT_CHECK_EQ(out->val.array.lparray[1].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[1].val.num, 2.5));
    RT_CHECK_EQ(out->val.array.lparray[2].xltype, static_cast<DWORD>(xltypeBool));
    RT_CHECK(out->val.array.lparray[2].val.xbool != 0);
    // Element strings allocated by GridToXLOPER12 carry xlbitDLLFree.
    RT_CHECK_EQ(out->val.array.lparray[3].xltype,
                static_cast<DWORD>(xltypeStr | xlbitDLLFree));
    RT_CHECK(out->val.array.lparray[3].val.str != nullptr);
    if (out->val.array.lparray[3].val.str) {
        RT_CHECK_EQ(static_cast<int>(out->val.array.lparray[3].val.str[0]), 2);
        RT_CHECK_EQ(out->val.array.lparray[3].val.str[1], L'o');
        RT_CHECK_EQ(out->val.array.lparray[3].val.str[2], L'k');
    }

    delete[] cells[3].val.str;
    xlAutoFree12(out);
}

static void TestGrid_1x1_Single() {
    XLOPER12 cell{};
    cell.xltype = xltypeNum;
    cell.val.num = 99.0;

    XLOPER12 multi{};
    multi.xltype = xltypeMulti;
    multi.val.array.rows = 1;
    multi.val.array.columns = 1;
    multi.val.array.lparray = &cell;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertGrid(&multi, builder);
    builder.Finish(off);
    auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 out = GridToXLOPER12(grid);

    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 1);
    RT_CHECK_EQ(out->val.array.columns, 1);
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[0].val.num, 99.0));
    xlAutoFree12(out);
}

static void TestGrid_Empty() {
    XLOPER12 multi{};
    multi.xltype = xltypeMulti;
    multi.val.array.rows = 0;
    multi.val.array.columns = 0;
    multi.val.array.lparray = nullptr;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertGrid(&multi, builder);
    builder.Finish(off);
    auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 out = GridToXLOPER12(grid);

    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 0);
    RT_CHECK_EQ(out->val.array.columns, 0);
    xlAutoFree12(out);
}

// ----------------------------------------------------------------------------
// NumGrid round-trips.

static void TestNumGrid_2x2() {
    FP12* fp = NewFP12(2, 2);
    fp->array[0] = 1.0;
    fp->array[1] = 2.0;
    fp->array[2] = 3.0;
    fp->array[3] = 4.0;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertNumGrid(fp, builder);
    builder.Finish(off);
    auto* ng = flatbuffers::GetRoot<protocol::NumGrid>(builder.GetBufferPointer());
    FP12* out = NumGridToFP12(ng);

    RT_CHECK(out != nullptr);
    if (out) {
        RT_CHECK_EQ(out->rows, 2);
        RT_CHECK_EQ(out->columns, 2);
        RT_CHECK(BitEqualDouble(out->array[0], 1.0));
        RT_CHECK(BitEqualDouble(out->array[1], 2.0));
        RT_CHECK(BitEqualDouble(out->array[2], 3.0));
        RT_CHECK(BitEqualDouble(out->array[3], 4.0));
    }
    // FP12 is allocated from a per-thread ring buffer (see mem.cpp's
    // NewFP12 / fpRingBuffers); no explicit free.
}

static void TestNumGrid_WithSpecials() {
    FP12* fp = NewFP12(1, 4);
    fp->array[0] = std::numeric_limits<double>::quiet_NaN();
    fp->array[1] = std::numeric_limits<double>::infinity();
    fp->array[2] = -std::numeric_limits<double>::infinity();
    fp->array[3] = -0.0;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertNumGrid(fp, builder);
    builder.Finish(off);
    auto* ng = flatbuffers::GetRoot<protocol::NumGrid>(builder.GetBufferPointer());
    FP12* out = NumGridToFP12(ng);

    RT_CHECK(out != nullptr);
    if (out) {
        RT_CHECK_EQ(out->rows, 1);
        RT_CHECK_EQ(out->columns, 4);
        RT_CHECK(std::isnan(out->array[0]));
        RT_CHECK(std::isinf(out->array[1]) && out->array[1] > 0);
        RT_CHECK(std::isinf(out->array[2]) && out->array[2] < 0);
        RT_CHECK(BitEqualDouble(out->array[3], -0.0));
    }
}

// ----------------------------------------------------------------------------
// Range round-trips.

static void TestRange_Sref_SingleRect() {
    XLOPER12 in{};
    in.xltype = xltypeSRef;
    in.val.sref.count = 1;
    in.val.sref.ref.rwFirst = 5;
    in.val.sref.ref.rwLast = 10;
    in.val.sref.ref.colFirst = 2;
    in.val.sref.ref.colLast = 7;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertRange(&in, builder);
    builder.Finish(off);
    auto* range = flatbuffers::GetRoot<protocol::Range>(builder.GetBufferPointer());
    LPXLOPER12 out = RangeToXLOPER12(range);

    // RangeToXLOPER12 always emits xltypeRef regardless of input being
    // SRef — that's by design (protocol::Range carries a list of rects, not
    // the single-ref optimization).
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeRef));
    RT_CHECK(out->val.mref.lpmref != nullptr);
    if (out->val.mref.lpmref) {
        RT_CHECK_EQ(static_cast<int>(out->val.mref.lpmref->count), 1);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].rwFirst, 5);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].rwLast, 10);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].colFirst, 2);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].colLast, 7);
    }
    xlAutoFree12(out);
}

// Sheet-name graceful degradation. ConvertRange now resolves Range.sheet_name
// via xlSheetNm/xlSheetId, but in the unit-test harness Excel is not bound
// (pexcel12 == NULL, SetExcel12EntryPt is never called), so every Excel12()
// call returns xlretFailed. The contract: the lookup must degrade to an EMPTY
// sheet_name without crashing, and the rects + format must survive intact.

static void TestRange_Sref_NoExcel_EmptySheetName() {
    XLOPER12 in{};
    in.xltype = xltypeSRef;
    in.val.sref.count = 1;
    in.val.sref.ref.rwFirst = 3;
    in.val.sref.ref.rwLast = 4;
    in.val.sref.ref.colFirst = 1;
    in.val.sref.ref.colLast = 2;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertRange(&in, builder, "General");
    builder.Finish(off);
    auto* range = flatbuffers::GetRoot<protocol::Range>(builder.GetBufferPointer());

    // No live Excel -> sheet_name omitted (null offset) => empty on read.
    RT_CHECK(range->sheet_name() == nullptr ||
             range->sheet_name()->size() == 0);
    // Rects intact.
    RT_CHECK(range->refs() != nullptr);
    if (range->refs()) {
        RT_CHECK_EQ(range->refs()->size(), 1u);
        const auto* r = range->refs()->Get(0);
        RT_CHECK_EQ(r->row_first(), 3);
        RT_CHECK_EQ(r->row_last(), 4);
        RT_CHECK_EQ(r->col_first(), 1);
        RT_CHECK_EQ(r->col_last(), 2);
    }
    // Format intact.
    RT_CHECK(range->format() != nullptr);
    if (range->format()) {
        RT_CHECK(range->format()->str() == "General");
    }
}

static void TestRange_Ref_NoExcel_EmptySheetName() {
    // Synthesize an xltypeRef with a single rect and a non-zero idSheet. Even
    // with a plausible idSheet, xlSheetNm is unreachable in tests, so the
    // name resolves empty and the rect survives.
    XLMREF12 mref{};
    mref.count = 1;
    mref.reftbl[0].rwFirst = 10;
    mref.reftbl[0].rwLast = 12;
    mref.reftbl[0].colFirst = 5;
    mref.reftbl[0].colLast = 6;

    XLOPER12 in{};
    in.xltype = xltypeRef;
    in.val.mref.lpmref = &mref;
    in.val.mref.idSheet = (IDSHEET)0x1234; // bogus but non-zero

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertRange(&in, builder, "0.00");
    builder.Finish(off);
    auto* range = flatbuffers::GetRoot<protocol::Range>(builder.GetBufferPointer());

    RT_CHECK(range->sheet_name() == nullptr ||
             range->sheet_name()->size() == 0);
    RT_CHECK(range->refs() != nullptr);
    if (range->refs()) {
        RT_CHECK_EQ(range->refs()->size(), 1u);
        const auto* r = range->refs()->Get(0);
        RT_CHECK_EQ(r->row_first(), 10);
        RT_CHECK_EQ(r->row_last(), 12);
        RT_CHECK_EQ(r->col_first(), 5);
        RT_CHECK_EQ(r->col_last(), 6);
    }
    RT_CHECK(range->format() != nullptr);
    if (range->format()) {
        RT_CHECK(range->format()->str() == "0.00");
    }
}

// ----------------------------------------------------------------------------
// Any-union round-trips: confirm ConvertAny routes each xltype to the right
// variant and AnyToXLOPER12 reconstructs it.

static void TestAny_Variant_Int() {
    XLOPER12 in{};
    in.xltype = xltypeInt;
    in.val.w = 7;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Int);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeInt));
    RT_CHECK_EQ(out->val.w, 7);
    xlAutoFree12(out);
}

static void TestAny_Variant_Num() {
    XLOPER12 in{};
    in.xltype = xltypeNum;
    in.val.num = 1.5;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Num);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.num, 1.5));
    xlAutoFree12(out);
}

static void TestAny_Variant_Bool() {
    XLOPER12 in{};
    in.xltype = xltypeBool;
    in.val.xbool = 1;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Bool);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeBool));
    RT_CHECK(out->val.xbool != 0);
    xlAutoFree12(out);
}

static void TestAny_Variant_Str() {
    XLOPER12 in{};
    in.xltype = xltypeStr;
    in.val.str = MakePascalString(L"any");

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Str);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeStr));
    if (out->val.str) {
        RT_CHECK_EQ(static_cast<int>(out->val.str[0]), 3);
    }
    delete[] in.val.str;
    xlAutoFree12(out);
}

static void TestAny_Variant_Err() {
    XLOPER12 in{};
    in.xltype = xltypeErr;
    in.val.err = xlerrNA;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Err);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeErr));
    RT_CHECK_EQ(out->val.err, xlerrNA);
    xlAutoFree12(out);
}

static void TestAny_Variant_Nil() {
    XLOPER12 in{};
    in.xltype = xltypeNil;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Nil);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNil));
    xlAutoFree12(out);
}

static void TestAny_Variant_Missing() {
    XLOPER12 in{};
    in.xltype = xltypeMissing;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    // xltypeMissing collapses to AnyValue::Nil in ConvertAny; check that
    // the resulting xltype is xltypeNil on the way back.
    RT_CHECK(any->val_type() == protocol::AnyValue::Nil);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeNil));
    xlAutoFree12(out);
}

static void TestAny_Variant_Grid_Mixed() {
    // Mixed-type 2x1 -> AnyValue::Grid (NOT NumGrid, since not homogenous).
    XLOPER12 cells[2];
    cells[0].xltype = xltypeNum;
    cells[0].val.num = 1.0;
    cells[1].xltype = xltypeStr;
    cells[1].val.str = MakePascalString(L"x");

    XLOPER12 multi{};
    multi.xltype = xltypeMulti;
    multi.val.array.rows = 2;
    multi.val.array.columns = 1;
    multi.val.array.lparray = cells;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&multi, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Grid);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 2);
    RT_CHECK_EQ(out->val.array.columns, 1);
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[0].val.num, 1.0));
    RT_CHECK_EQ(out->val.array.lparray[1].xltype,
                static_cast<DWORD>(xltypeStr | xlbitDLLFree));
    delete[] cells[1].val.str;
    xlAutoFree12(out);
}

static void TestAny_Variant_NumGrid_Homogenous() {
    // Homogenous-num 2x1 -> AnyValue::NumGrid (ConvertMultiToAny shortcut).
    XLOPER12 cells[2];
    cells[0].xltype = xltypeNum;
    cells[0].val.num = 10.0;
    cells[1].xltype = xltypeNum;
    cells[1].val.num = 20.0;

    XLOPER12 multi{};
    multi.xltype = xltypeMulti;
    multi.val.array.rows = 2;
    multi.val.array.columns = 1;
    multi.val.array.lparray = cells;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&multi, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::NumGrid);

    LPXLOPER12 out = AnyToXLOPER12(any);
    // AnyToXLOPER12 expands NumGrid into a Multi of xltypeNum cells.
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.rows, 2);
    RT_CHECK_EQ(out->val.array.columns, 1);
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[0].val.num, 10.0));
    RT_CHECK_EQ(out->val.array.lparray[1].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK(BitEqualDouble(out->val.array.lparray[1].val.num, 20.0));
    xlAutoFree12(out);
}

static void TestAny_Variant_Range() {
    XLOPER12 in{};
    in.xltype = xltypeSRef;
    in.val.sref.count = 1;
    in.val.sref.ref.rwFirst = 0;
    in.val.sref.ref.rwLast = 1;
    in.val.sref.ref.colFirst = 2;
    in.val.sref.ref.colLast = 3;

    flatbuffers::FlatBufferBuilder builder;
    auto off = ConvertAny(&in, builder);
    builder.Finish(off);
    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Range);

    LPXLOPER12 out = AnyToXLOPER12(any);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeRef));
    RT_CHECK(out->val.mref.lpmref != nullptr);
    if (out->val.mref.lpmref) {
        RT_CHECK_EQ(static_cast<int>(out->val.mref.lpmref->count), 1);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].rwFirst, 0);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].rwLast, 1);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].colFirst, 2);
        RT_CHECK_EQ(out->val.mref.lpmref->reftbl[0].colLast, 3);
    }
    xlAutoFree12(out);
}

// ----------------------------------------------------------------------------
// xltypeMulti element-string ownership (xlbitDLLFree contract).
//
// GridToXLOPER12 marks every element string it allocates with xlbitDLLFree;
// xlAutoFree12 (and GridToXLOPER12's internal ScopeGuard) only delete[]
// element strings carrying that bit. Elements without the bit are borrowed
// and must be left alone.

static void TestMulti_OwnedStrings_CarryDLLFreeBit() {
    // Mixed grid (num + string) built by GridToXLOPER12: the string element
    // must carry xlbitDLLFree, the num element must not. Freeing via
    // xlAutoFree12 must reclaim the marked string and the array without
    // crashing.
    flatbuffers::FlatBufferBuilder builder;
    std::vector<flatbuffers::Offset<protocol::Scalar>> elems;
    elems.push_back(protocol::CreateScalar(
        builder, protocol::ScalarValue::Num,
        protocol::CreateNum(builder, 1.5).Union()));
    elems.push_back(protocol::CreateScalar(
        builder, protocol::ScalarValue::Str,
        protocol::CreateStr(builder, builder.CreateString("owned")).Union()));
    auto vec = builder.CreateVector(elems);
    auto gridOff = protocol::CreateGrid(builder, 1, 2, vec);
    builder.Finish(gridOff);
    auto* grid = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());

    LPXLOPER12 out = GridToXLOPER12(grid);
    RT_CHECK_EQ(CoreType(*out), static_cast<DWORD>(xltypeMulti));
    RT_CHECK_EQ(out->val.array.lparray[0].xltype, static_cast<DWORD>(xltypeNum));
    RT_CHECK_EQ(out->val.array.lparray[1].xltype,
                static_cast<DWORD>(xltypeStr | xlbitDLLFree));
    RT_CHECK(out->val.array.lparray[1].val.str != nullptr);
    xlAutoFree12(out);
}

static void TestMulti_NonOwnedString_NotDeleted() {
    // Negative test: an element string WITHOUT xlbitDLLFree (simulating an
    // Excel-owned/aliased pointer placed into a multi cell) must NOT be
    // delete[]'d by xlAutoFree12. Under the pre-contract implementation this
    // called delete[] on a static buffer -> heap corruption/abort.
    static XCHAR staticStr[] = { 2, L'h', L'i', 0 };

    LPXLOPER12 op = NewXLOPER12();
    op->xltype = xltypeMulti | xlbitDLLFree;
    op->val.array.rows = 1;
    op->val.array.columns = 2;
    op->val.array.lparray = new XLOPER12[2];
    std::memset(op->val.array.lparray, 0, 2 * sizeof(XLOPER12));

    // Element 0: borrowed string — no ownership bit.
    op->val.array.lparray[0].xltype = xltypeStr;
    op->val.array.lparray[0].val.str = staticStr;

    // Element 1: DLL-owned string, allocated and marked the same way
    // GridToXLOPER12 does. Must be freed (leak-checked under ASan/CRT).
    XCHAR* owned = new XCHAR[4];
    owned[0] = 2; owned[1] = L'o'; owned[2] = L'k'; owned[3] = 0;
    op->val.array.lparray[1].xltype = xltypeStr | xlbitDLLFree;
    op->val.array.lparray[1].val.str = owned;

    xlAutoFree12(op);  // must skip element 0, free element 1 + array

    // If the static buffer had been delete[]'d the runtime would typically
    // have aborted above; verify its contents survived as a best-effort
    // observable check.
    RT_CHECK_EQ(static_cast<int>(staticStr[0]), 2);
    RT_CHECK_EQ(staticStr[1], L'h');
    RT_CHECK_EQ(staticStr[2], L'i');
}

static void TestMulti_EchoRoundTrip_MasksOwnershipBits() {
    // Masking audit regression: a multi produced by GridToXLOPER12 (string
    // elements carry xlbitDLLFree) fed back through the Excel->FlatBuffers
    // direction must classify element types correctly. Without masking in
    // ConvertScalar/ConvertMultiToAny the string element would silently
    // degrade to Nil and the mixed grid could be misrouted.
    flatbuffers::FlatBufferBuilder b1;
    std::vector<flatbuffers::Offset<protocol::Scalar>> elems;
    elems.push_back(protocol::CreateScalar(
        b1, protocol::ScalarValue::Num, protocol::CreateNum(b1, 2.5).Union()));
    elems.push_back(protocol::CreateScalar(
        b1, protocol::ScalarValue::Str,
        protocol::CreateStr(b1, b1.CreateString("echo")).Union()));
    auto vec1 = b1.CreateVector(elems);
    auto gridOff = protocol::CreateGrid(b1, 1, 2, vec1);
    b1.Finish(gridOff);
    auto* grid1 = flatbuffers::GetRoot<protocol::Grid>(b1.GetBufferPointer());

    LPXLOPER12 mid = GridToXLOPER12(grid1);
    RT_CHECK_EQ(CoreType(*mid), static_cast<DWORD>(xltypeMulti));

    // Echo leg 1: ConvertGrid -> element types must survive.
    flatbuffers::FlatBufferBuilder b2;
    auto off2 = ConvertGrid(mid, b2);
    b2.Finish(off2);
    auto* grid2 = flatbuffers::GetRoot<protocol::Grid>(b2.GetBufferPointer());
    RT_CHECK_EQ(grid2->rows(), 1);
    RT_CHECK_EQ(grid2->cols(), 2);
    RT_CHECK(grid2->data()->Get(0)->val_type() == protocol::ScalarValue::Num);
    RT_CHECK(grid2->data()->Get(1)->val_type() == protocol::ScalarValue::Str);
    if (grid2->data()->Get(1)->val_type() == protocol::ScalarValue::Str) {
        RT_CHECK(grid2->data()->Get(1)->val_as_Str()->val()->str() == "echo");
    }

    // Echo leg 2: ConvertAny on the (mixed) multi must route to Grid, not
    // misclassify because of the per-element ownership bits.
    flatbuffers::FlatBufferBuilder b3;
    auto off3 = ConvertAny(mid, b3);
    b3.Finish(off3);
    auto* any = flatbuffers::GetRoot<protocol::Any>(b3.GetBufferPointer());
    RT_CHECK(any->val_type() == protocol::AnyValue::Grid);

    xlAutoFree12(mid);
}

// ----------------------------------------------------------------------------

int main() {
    std::cout << "Running XLOPER12 <-> FlatBuffer round-trip suite" << std::endl;

    // Scalar (per type, with edge cases)
    RT_RUN(TestScalar_Int_Basic);
    RT_RUN(TestScalar_Int_Negative);
    RT_RUN(TestScalar_Int_Min);
    RT_RUN(TestScalar_Int_Max);
    RT_RUN(TestScalar_Num_Basic);
    RT_RUN(TestScalar_Num_Zero);
    RT_RUN(TestScalar_Num_NegativeZero);
    RT_RUN(TestScalar_Num_NaN);
    RT_RUN(TestScalar_Num_PosInf);
    RT_RUN(TestScalar_Num_NegInf);
    RT_RUN(TestScalar_Str_Ascii);
    RT_RUN(TestScalar_Str_Empty);
    RT_RUN(TestScalar_Str_Unicode);
    RT_RUN(TestScalar_Bool_True);
    RT_RUN(TestScalar_Bool_False);
    RT_RUN(TestScalar_Err_Value);
    RT_RUN(TestScalar_Err_Num);
    RT_RUN(TestScalar_Err_Div0);
    RT_RUN(TestConvertScalar_Direct_Num);

    // Grid
    RT_RUN(TestGrid_2x2_Mixed);
    RT_RUN(TestGrid_1x1_Single);
    RT_RUN(TestGrid_Empty);

    // NumGrid
    RT_RUN(TestNumGrid_2x2);
    RT_RUN(TestNumGrid_WithSpecials);

    // Range
    RT_RUN(TestRange_Sref_SingleRect);
    RT_RUN(TestRange_Sref_NoExcel_EmptySheetName);
    RT_RUN(TestRange_Ref_NoExcel_EmptySheetName);

    // Any union variants
    RT_RUN(TestAny_Variant_Int);
    RT_RUN(TestAny_Variant_Num);
    RT_RUN(TestAny_Variant_Bool);
    RT_RUN(TestAny_Variant_Str);
    RT_RUN(TestAny_Variant_Err);
    RT_RUN(TestAny_Variant_Nil);
    RT_RUN(TestAny_Variant_Missing);
    RT_RUN(TestAny_Variant_Grid_Mixed);
    RT_RUN(TestAny_Variant_NumGrid_Homogenous);
    RT_RUN(TestAny_Variant_Range);

    // xltypeMulti element-string ownership (xlbitDLLFree contract)
    RT_RUN(TestMulti_OwnedStrings_CarryDLLFreeBit);
    RT_RUN(TestMulti_NonOwnedString_NotDeleted);
    RT_RUN(TestMulti_EchoRoundTrip_MasksOwnershipBits);

    if (g_failures == 0) {
        std::cout << "All round-trip tests passed." << std::endl;
        return 0;
    }
    std::cerr << g_failures << " failure(s)." << std::endl;
    return 1;
}
