#include <cassert>
#include <cstring>
#include <iostream>
#include <string>

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

void TestStrConversion() {
    flatbuffers::FlatBufferBuilder builder;
    XLOPER12 op;
    op.xltype = xltypeStr;
    std::wstring s = L"Hello World";

    // Manually create a pascal string for input
    op.val.str = new XCHAR[s.length() + 1];
    op.val.str[0] = (XCHAR)s.length();
    for (size_t i = 0; i < s.length(); ++i) op.val.str[i + 1] = s[i];

    auto offset = ConvertAny(&op, builder);
    builder.Finish(offset);

    auto* any = flatbuffers::GetRoot<protocol::Any>(builder.GetBufferPointer());
    assert(any->val_type() == protocol::AnyValue::Str);
    assert(any->val_as_Str()->val()->str() == "Hello World");

    LPXLOPER12 res = AnyToXLOPER12(any);
    assert(res->xltype == (xltypeStr | xlbitDLLFree));
    assert(res->val.str[0] == 11);
    for (size_t i = 0; i < 11; ++i) assert(res->val.str[i + 1] == s[i]);

    delete[] op.val.str;
    xlAutoFree12(res);
    std::cout << "TestStrConversion passed" << std::endl;
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
    TestStrConversion();
    TestNilConversion();
    std::cout << "All tests passed!" << std::endl;
    return 0;
}
