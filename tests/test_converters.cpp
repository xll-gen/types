#include <iostream>
#include <cassert>
#include <cstring>
#include <string>
#include <vector>

#include <windows.h>
HINSTANCE g_hModule = NULL;

#include "types/converters.h"
#include "types/mem.h"

extern "C" void __stdcall xlAutoFree12(LPXLOPER12 p);

void TestNumConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeNum;
    op.val.num = 123.456;

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Num);
    assert(any->val_as_Num()->val() == 123.456);

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeNum | xlbitDLLFree));
    assert(res->val.num == 123.456);

    xlAutoFree12(res);
    std::cout << "TestNumConversion passed" << std::endl;
}

void TestIntConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeInt;
    op.val.w = 42;

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Int);
    assert(any->val_as_Int()->val() == 42);

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeInt | xlbitDLLFree));
    assert(res->val.w == 42);

    xlAutoFree12(res);
    std::cout << "TestIntConversion passed" << std::endl;
}

void TestBoolConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeBool;
    op.val.xbool = 1;

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Bool);
    assert(any->val_as_Bool()->val() == true);

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeBool | xlbitDLLFree));
    assert(res->val.xbool == 1);

    xlAutoFree12(res);
    std::cout << "TestBoolConversion passed" << std::endl;
}

void TestErrConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeErr;
    op.val.err = 15; // xlerrValue

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Err);
    assert(any->val_as_Err()->val() == protocol::XlError::Value); // 2015

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeErr | xlbitDLLFree));
    assert(res->val.err == 15);

    xlAutoFree12(res);
    std::cout << "TestErrConversion passed" << std::endl;
}

void TestStrConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeStr;
    std::wstring s = L"Hello World";

    // Manually create a pascal string for input
    op.val.str = new XCHAR[s.length() + 1];
    op.val.str[0] = (XCHAR)s.length();
    for(size_t i=0; i<s.length(); ++i) op.val.str[i+1] = s[i];

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Str);
    assert(any->val_as_Str()->val()->str() == "Hello World");

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeStr | xlbitDLLFree));
    assert(res->val.str[0] == 11);
    for(size_t i=0; i<11; ++i) assert(res->val.str[i+1] == s[i]);

    delete[] op.val.str;
    xlAutoFree12(res);
    std::cout << "TestStrConversion passed" << std::endl;
}

void TestGridConversion() {
    flatbuffers::FlatBufferBuilder builder;
    std::vector<flatbuffers::Offset<protocol::Scalar>> elements;

    // Mixed content: Num, Str
    elements.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Num, protocol::CreateNum(builder, 1.23).Union()));
    elements.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Str, protocol::CreateStr(builder, builder.CreateString("Test")).Union()));

    auto vec = builder.CreateVector(elements);
    auto grid = protocol::CreateGrid(builder, 2, 1, vec); // 2 Rows, 1 Col
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Grid, grid.Union());
    builder.Finish(any);

    auto* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    LPXLOPER12 res = AnyToXLOPER12(root);

    assert(res->xltype == (xltypeMulti | xlbitDLLFree));
    assert(res->val.array.rows == 2);
    assert(res->val.array.columns == 1);

    assert(res->val.array.lparray[0].xltype == xltypeNum);
    assert(res->val.array.lparray[0].val.num == 1.23);

    // String elements we allocate carry xlbitDLLFree (ownership marker for
    // xlAutoFree12); non-heap elements (Num above) stay unmarked.
    assert(res->val.array.lparray[1].xltype == (xltypeStr | xlbitDLLFree));
    assert(res->val.array.lparray[1].val.str[0] == 4);

    xlAutoFree12(res);
    std::cout << "TestGridConversion passed" << std::endl;
}

void TestNumGridConversion() {
    flatbuffers::FlatBufferBuilder builder;
    std::vector<double> data = { 10.0, 20.0, 30.0, 40.0 };
    auto vec = builder.CreateVector(data);
    auto ng = protocol::CreateNumGrid(builder, 2, 2, vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
    builder.Finish(any);

    auto* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    LPXLOPER12 res = AnyToXLOPER12(root);

    assert(res->xltype == (xltypeMulti | xlbitDLLFree)); // NumGrid converts to Multi
    assert(res->val.array.rows == 2);
    assert(res->val.array.columns == 2);
    assert(res->val.array.lparray[0].val.num == 10.0);
    assert(res->val.array.lparray[3].val.num == 40.0);

    xlAutoFree12(res);
    std::cout << "TestNumGridConversion passed" << std::endl;
}

void TestRangeConversion() {
    flatbuffers::FlatBufferBuilder builder;
    std::vector<protocol::Rect> rects;
    rects.emplace_back(1, 2, 3, 4); // R1:R2, C3:C4

    auto vec = builder.CreateVectorOfStructs(rects);
    auto rng = protocol::CreateRange(builder, 0, vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Range, rng.Union());
    builder.Finish(any);

    auto* root = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    LPXLOPER12 res = AnyToXLOPER12(root);

    assert(res->xltype == (xltypeRef | xlbitDLLFree));
    assert(res->val.mref.lpmref->count == 1);
    assert(res->val.mref.lpmref->reftbl[0].rwFirst == 1);
    assert(res->val.mref.lpmref->reftbl[0].rwLast == 2);
    assert(res->val.mref.lpmref->reftbl[0].colFirst == 3);
    assert(res->val.mref.lpmref->reftbl[0].colLast == 4);

    xlAutoFree12(res);
    std::cout << "TestRangeConversion passed" << std::endl;
}

void TestAnyDateBecomesNum() {
    flatbuffers::FlatBufferBuilder builder;
    auto d = protocol::CreateDate(builder, 46188.5 /*serial*/, 0 /*format*/);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Date, d.Union());
    builder.Finish(any);

    auto* a = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(a->val_type() == protocol::AnyValue::Date);

    LPXLOPER12 op = AnyToXLOPER12(a);
    assert(op != nullptr);
    assert((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == (DWORD)xltypeNum);
    assert(op->val.num == 46188.5);

    xlAutoFree12(op);
    std::cout << "TestAnyDateBecomesNum passed" << std::endl;
}

void TestGridDateCellBecomesNum() {
    flatbuffers::FlatBufferBuilder builder;
    auto d = protocol::CreateDate(builder, 46188.0, 0);
    auto cell = protocol::CreateScalar(builder, protocol::ScalarValue::Date, d.Union());
    std::vector<flatbuffers::Offset<protocol::Scalar>> cells{cell};
    auto vec = builder.CreateVector(cells);
    auto grid = protocol::CreateGrid(builder, 1, 1, vec);
    builder.Finish(grid);

    auto* g = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    LPXLOPER12 op = GridToXLOPER12(g);
    assert(op != nullptr);
    assert((op->xltype & ~(xlbitDLLFree | xlbitXLFree)) == (DWORD)xltypeMulti);
    assert(op->val.array.lparray[0].xltype == (DWORD)xltypeNum);
    assert(op->val.array.lparray[0].val.num == 46188.0);

    xlAutoFree12(op);
    std::cout << "TestGridDateCellBecomesNum passed" << std::endl;
}

void TestCollectDateCells_ScalarDate() {
    flatbuffers::FlatBufferBuilder builder;
    // serial 46188.0 is integer (no fractional part), format index 0 (none).
    auto d = protocol::CreateDate(builder, 46188.0, 0);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Date, d.Union());
    builder.Finish(any);

    auto* a = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    std::vector<DateCell> cells;
    CollectDateCells(a, cells);

    assert(cells.size() == 1);
    assert(cells[0].rowOff == 0);
    assert(cells[0].colOff == 0);
    assert(cells[0].format == L"yyyy-mm-dd");
    std::cout << "TestCollectDateCells_ScalarDate passed" << std::endl;
}

void TestCollectDateCells_GridColumnZeroOnly() {
    flatbuffers::FlatBufferBuilder builder;
    // 2x2 grid (row-major): col0 = Date cells, col1 = Num.
    // Layout indices: [0]=date 46188, [1]=num 1.5, [2]=date 46189, [3]=num 1.5
    std::vector<flatbuffers::Offset<protocol::Scalar>> cells;
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Date,
                                           protocol::CreateDate(builder, 46188.0, 0).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Num,
                                           protocol::CreateNum(builder, 1.5).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Date,
                                           protocol::CreateDate(builder, 46189.0, 0).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Num,
                                           protocol::CreateNum(builder, 1.5).Union()));
    auto vec = builder.CreateVector(cells);
    auto grid = protocol::CreateGrid(builder, 2, 2, vec); // 2 rows, 2 cols
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Grid, grid.Union());
    builder.Finish(any);

    auto* a = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    std::vector<DateCell> out;
    CollectDateCells(a, out);

    // Only the date column (col0) yields cells: rows 0 and 1, both colOff 0.
    assert(out.size() == 2);
    assert(out[0].rowOff == 0 && out[0].colOff == 0);
    assert(out[1].rowOff == 1 && out[1].colOff == 0);
    std::cout << "TestCollectDateCells_GridColumnZeroOnly passed" << std::endl;
}

void TestCollectDateCells_DatetimeUsesDatetimeFormat() {
    flatbuffers::FlatBufferBuilder builder;
    // serial 46188.5 has a fractional part -> datetime auto-format.
    auto d = protocol::CreateDate(builder, 46188.5, 0);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::Date, d.Union());
    builder.Finish(any);

    auto* a = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    std::vector<DateCell> cells;
    CollectDateCells(a, cells);

    assert(cells.size() == 1);
    assert(cells[0].format == L"yyyy-mm-dd hh:mm:ss");
    std::cout << "TestCollectDateCells_DatetimeUsesDatetimeFormat passed" << std::endl;
}

void TestCollectDateCells_NumGridHasNoDates() {
    flatbuffers::FlatBufferBuilder builder;
    std::vector<double> data = { 10.0, 20.0, 30.0, 40.0 };
    auto vec = builder.CreateVector(data);
    auto ng = protocol::CreateNumGrid(builder, 2, 2, vec);
    auto any = protocol::CreateAny(builder, protocol::AnyValue::NumGrid, ng.Union());
    builder.Finish(any);

    auto* a = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    std::vector<DateCell> cells;
    CollectDateCells(a, cells);

    assert(cells.empty());
    std::cout << "TestCollectDateCells_NumGridHasNoDates passed" << std::endl;
}

void TestCollectDateCells_GridOverloadDirect() {
    // Same 2x2 grid as TestCollectDateCells_GridColumnZeroOnly, but built as a
    // bare Grid (no Any wrapper) and passed to the Grid overload directly —
    // mirrors the sync grid-return wrapper path.
    flatbuffers::FlatBufferBuilder builder;
    std::vector<flatbuffers::Offset<protocol::Scalar>> cells;
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Date,
                                           protocol::CreateDate(builder, 46188.0, 0).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Num,
                                           protocol::CreateNum(builder, 1.5).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Date,
                                           protocol::CreateDate(builder, 46189.0, 0).Union()));
    cells.push_back(protocol::CreateScalar(builder, protocol::ScalarValue::Num,
                                           protocol::CreateNum(builder, 1.5).Union()));
    auto vec = builder.CreateVector(cells);
    auto grid = protocol::CreateGrid(builder, 2, 2, vec);
    builder.Finish(grid);

    auto* g = flatbuffers::GetRoot<protocol::Grid>(builder.GetBufferPointer());
    std::vector<DateCell> out;
    CollectDateCells(g, out);

    assert(out.size() == 2);
    assert(out[0].rowOff == 0 && out[0].colOff == 0);
    assert(out[1].rowOff == 1 && out[1].colOff == 0);

    std::vector<DateCell> none;
    CollectDateCells((const protocol::Grid*)nullptr, none);
    assert(none.empty());
    std::cout << "TestCollectDateCells_GridOverloadDirect passed" << std::endl;
}

void TestNilConversion() {
    // Test converting nil/missing
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeMissing;

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Nil);

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeNil | xlbitDLLFree));

    xlAutoFree12(res);

    // Test converting NULL any
    LPXLOPER12 resNull = AnyToXLOPER12(nullptr);
    assert(resNull->xltype == (xltypeNil | xlbitDLLFree));
    xlAutoFree12(resNull);

    std::cout << "TestNilConversion passed" << std::endl;
}

int main() {
    TestNumConversion();
    TestIntConversion();
    TestBoolConversion();
    TestErrConversion();
    TestStrConversion();
    TestGridConversion();
    TestNumGridConversion();
    TestRangeConversion();
    TestAnyDateBecomesNum();
    TestGridDateCellBecomesNum();
    TestCollectDateCells_ScalarDate();
    TestCollectDateCells_GridColumnZeroOnly();
    TestCollectDateCells_DatetimeUsesDatetimeFormat();
    TestCollectDateCells_NumGridHasNoDates();
    TestCollectDateCells_GridOverloadDirect();
    TestNilConversion();
    std::cout << "All tests passed!" << std::endl;
    return 0;
}
