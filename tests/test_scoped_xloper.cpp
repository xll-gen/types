#include <iostream>
#include <cassert>
#include <cstring>
#include <string>
#include <vector>
#include <cwchar>

#include "types/ScopedXLOPER12.h"

void TestIntConstructor() {
    ScopedXLOPER12 op(42);
    LPXLOPER12 p = op;
    assert(p->xltype == xltypeInt);
    assert(p->val.w == 42);
    std::cout << "TestIntConstructor passed" << std::endl;
}

void TestDoubleConstructor() {
    ScopedXLOPER12 op(3.14);
    LPXLOPER12 p = op;
    assert(p->xltype == xltypeNum);
    assert(p->val.num == 3.14);
    std::cout << "TestDoubleConstructor passed" << std::endl;
}

void TestBoolConstructor() {
    ScopedXLOPER12 op(true);
    LPXLOPER12 p = op;
    assert(p->xltype == xltypeBool);
    assert(p->val.xbool == 1);

    ScopedXLOPER12 op2(false);
    p = op2;
    assert(p->val.xbool == 0);
    std::cout << "TestBoolConstructor passed" << std::endl;
}

void TestStringConstructor() {
    const wchar_t* txt = L"Hello";
    ScopedXLOPER12 op(txt);
    LPXLOPER12 p = op;
    assert(p->xltype == xltypeStr);
    assert((size_t)p->val.str[0] == 5);
    assert(std::wcsncmp(p->val.str + 1, txt, 5) == 0);
    std::cout << "TestStringConstructor passed" << std::endl;
}

void TestStringTruncation() {
    std::wstring longStr(33000, L'a');
    ScopedXLOPER12 op(longStr);
    LPXLOPER12 p = op;
    assert(p->xltype == xltypeStr);
    assert((size_t)p->val.str[0] == 32767);
    std::cout << "TestStringTruncation passed" << std::endl;
}

void TestMoveConstructor() {
    ScopedXLOPER12 op(L"Moved");
    const wchar_t* originalPtr = op.get()->val.str;

    ScopedXLOPER12 op2(std::move(op));
    LPXLOPER12 p = op2;
    assert(p->xltype == xltypeStr);
    assert((size_t)p->val.str[0] == 5);
    assert(p->val.str == originalPtr);

    // Check old one is nil
    LPXLOPER12 oldP = op;
    assert(oldP->xltype == xltypeNil);

    std::cout << "TestMoveConstructor passed" << std::endl;
}

void TestVectorUsage() {
    std::vector<ScopedXLOPER12> args;
    args.emplace_back(10);
    args.emplace_back(L"Test");

    // Check if pointers are valid
    LPXLOPER12 p1 = args[0];
    assert(p1->xltype == xltypeInt);
    assert(p1->val.w == 10);

    LPXLOPER12 p2 = args[1];
    assert(p2->xltype == xltypeStr);
    assert((size_t)p2->val.str[0] == 4);
    assert(std::wcsncmp(p2->val.str + 1, L"Test", 4) == 0);
    std::cout << "TestVectorUsage passed" << std::endl;
}

// --- ScopedXLOPER12Result Tests ---

// Mock Excel12 entry point
// Needs to match internal typedef signature
typedef int (PASCAL *EXCEL12PROC) (int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res);

extern "C" void pascal SetExcel12EntryPt(EXCEL12PROC pexcel12New);

static bool g_xlFreeCalled = false;
static LPXLOPER12 g_freedOp = nullptr;

int PASCAL MockExcel12(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res) {
    if (xlfn == xlFree) {
        g_xlFreeCalled = true;
        // xlFree is called as Excel12(xlFree, 0, 1, &op)
        // coper should be 1
        // rgpxloper12[0] should be pointer to the op to free
        if (coper >= 1 && rgpxloper12) {
             g_freedOp = rgpxloper12[0];
        }
    }
    return xlretSuccess;
}

void TestResultAutoFree() {
    // Install mock
    SetExcel12EntryPt(MockExcel12);

    {
        ScopedXLOPER12Result res;
        LPXLOPER12 op = res;

        // Simulate a result with xlbitXLFree
        op->xltype = xltypeStr | xlbitXLFree;
        op->val.str = nullptr; // Dummy

        g_xlFreeCalled = false;
        g_freedOp = nullptr;
        // ScopedXLOPER12Result goes out of scope here
    }

    assert(g_xlFreeCalled);
    // Verify it tried to free the correct address
    // But 'res' is destroyed, so we can't check 'res' address unless we captured it.

    // Do it again capturing address
    g_xlFreeCalled = false;
    LPXLOPER12 capturedAddr = nullptr;
    {
        ScopedXLOPER12Result res;
        capturedAddr = res;
        res->xltype = xltypeNum | xlbitXLFree;
    }
    assert(g_xlFreeCalled);
    assert(g_freedOp == capturedAddr);

    std::cout << "TestResultAutoFree passed" << std::endl;
}

void TestResultNoFree() {
    // Install mock
    SetExcel12EntryPt(MockExcel12);

    {
        ScopedXLOPER12Result res;
        LPXLOPER12 op = res;

        // Result without xlbitXLFree (e.g. xltypeInt)
        op->xltype = xltypeInt;
        op->val.w = 123;

        g_xlFreeCalled = false;
    }

    assert(!g_xlFreeCalled);
    std::cout << "TestResultNoFree passed" << std::endl;
}

int main() {
    TestIntConstructor();
    TestDoubleConstructor();
    TestBoolConstructor();
    TestStringConstructor();
    TestStringTruncation();
    TestMoveConstructor();
    TestVectorUsage();

    TestResultAutoFree();
    TestResultNoFree();

    return 0;
}
