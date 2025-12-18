#include <iostream>
#include <cassert>
#include <cstring>
#include <string>
#include <vector>
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
    op.val.err = (int)protocol::XlError::Value;

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Err);
    assert(any->val_as_Err()->val() == protocol::XlError::Value);

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeErr | xlbitDLLFree));
    assert(res->val.err == (int)protocol::XlError::Value);

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

    assert(res->val.array.lparray[1].xltype == xltypeStr);
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

void TestNilConversion() {
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
    TestNilConversion();
    std::cout << "All tests passed!" << std::endl;
    return 0;
}
