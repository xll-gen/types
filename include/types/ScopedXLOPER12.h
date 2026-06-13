#pragma once

#include <windows.h>

#include "types/xlcall.h"
#include <vector>
#include <string>
#include <cstring>
#include <cwchar>
#include <utility> // for std::move

// A helper class to manage XLOPER12 memory for arguments passed to Excel.
// This is safer than using TempStr12/TempInt12 for generic wrappers because it
// manages memory lifetime explicitly and avoids ring buffer limits.
class ScopedXLOPER12 {
public:
    ScopedXLOPER12() {
        m_op.xltype = xltypeNil;
    }

    // Move constructor
    ScopedXLOPER12(ScopedXLOPER12&& other) noexcept : m_op(other.m_op), m_buffer(std::move(other.m_buffer)) {
        // If we moved a string, we must update the pointer to the new buffer location
        // because std::vector move might transfer ownership but the pointer value in m_op
        // needs to be valid.
        if (m_op.xltype == xltypeStr) {
            m_op.val.str = m_buffer.data();
        }
        other.m_op.xltype = xltypeNil;
    }

    // Delete copy to prevent accidental deep copies
    ScopedXLOPER12(const ScopedXLOPER12&) = delete;
    ScopedXLOPER12& operator=(const ScopedXLOPER12&) = delete;

    ScopedXLOPER12& operator=(ScopedXLOPER12&& other) noexcept {
        if (this != &other) {
            m_op = other.m_op;
            m_buffer = std::move(other.m_buffer);
            if (m_op.xltype == xltypeStr) {
                m_op.val.str = m_buffer.data();
            }
            other.m_op.xltype = xltypeNil;
        }
        return *this;
    }

    // Constructors for various types

    explicit ScopedXLOPER12(int val) {
        m_op.xltype = xltypeInt;
        m_op.val.w = val;
    }

    explicit ScopedXLOPER12(double val) {
        m_op.xltype = xltypeNum;
        m_op.val.num = val;
    }

    explicit ScopedXLOPER12(bool val) {
        m_op.xltype = xltypeBool;
        m_op.val.xbool = val ? 1 : 0;
    }

    explicit ScopedXLOPER12(const wchar_t* str) {
        SetString(str);
    }

    explicit ScopedXLOPER12(const std::wstring& str) {
        SetString(str.c_str());
    }

    // Support constructing from an existing XLOPER12 (shallow copy for scalars, deep for str)
    explicit ScopedXLOPER12(const XLOPER12* op) {
        if (!op) {
            m_op.xltype = xltypeNil;
            return;
        }
        if (op->xltype == xltypeStr) {
            if (!op->val.str) {
                // Malformed input (xltypeStr with null payload): degrade
                // gracefully instead of dereferencing null.
                m_op.xltype = xltypeMissing;
                return;
            }
            // Pascal string in op->val.str
            size_t len = (size_t)op->val.str[0];
            if (len > 32767) len = 32767;
            m_buffer.resize(len + 2); // +1 for length, +1 for safety
            std::memcpy(m_buffer.data(), op->val.str, (len + 1) * sizeof(wchar_t));
            m_op.xltype = xltypeStr;
            m_op.val.str = m_buffer.data();
        } else {
            m_op = *op;
        }
    }

    // Implicit conversion to LPXLOPER12 for easy passing to Excel12v
    operator LPXLOPER12() {
        return &m_op;
    }

    // Const version
    operator const XLOPER12*() const {
        return &m_op;
    }

    LPXLOPER12 get() {
        return &m_op;
    }

private:
    void SetString(const wchar_t* str) {
        if (!str) {
            m_op.xltype = xltypeMissing;
            return;
        }
        size_t len = std::wcslen(str);
        if (len > 32767) len = 32767; // Excel limit

        m_buffer.resize(len + 2); // 1 for length char, 1 for null term
        m_buffer[0] = (wchar_t)len;
        if (len > 0) {
            std::wmemcpy(m_buffer.data() + 1, str, len);
        }
        m_buffer[len + 1] = 0; // Null terminate

        m_op.xltype = xltypeStr;
        m_op.val.str = m_buffer.data();
    }

    XLOPER12 m_op;
    std::vector<wchar_t> m_buffer; // Used for xltypeStr
};

// A helper class to manage the result XLOPER12 from Excel callbacks.
// The destructor releases Excel-allocated payloads (str/multi/ref) via xlFree.
// INVARIANT: only ever store Excel12()/Excel12v() RESULT operands here. Never
// put a DLL-allocated (xlbitDLLFree-tagged, ObjectPool/NewXLOPER12) value in a
// Result — those must go through xlAutoFree12, not xlFree.
class ScopedXLOPER12Result {
public:
    ScopedXLOPER12Result() {
        m_op.xltype = xltypeNil;
    }

    ~ScopedXLOPER12Result() {
        Free();
    }

    // No copy
    ScopedXLOPER12Result(const ScopedXLOPER12Result&) = delete;
    ScopedXLOPER12Result& operator=(const ScopedXLOPER12Result&) = delete;

    // Move
    ScopedXLOPER12Result(ScopedXLOPER12Result&& other) noexcept : m_op(other.m_op) {
        other.m_op.xltype = xltypeNil;
    }

    ScopedXLOPER12Result& operator=(ScopedXLOPER12Result&& other) noexcept {
        if (this != &other) {
            Free();
            m_op = other.m_op;
            other.m_op.xltype = xltypeNil;
        }
        return *this;
    }

    // Accessors
    LPXLOPER12 get() { return &m_op; }
    operator LPXLOPER12() { return &m_op; }
    LPXLOPER12 operator->() { return &m_op; }

private:
    void Free() {
        // This class holds results populated by Excel12()/Excel12v(), and the
        // SDK contract is that the CALLER releases such results with xlFree.
        // Excel never sets xlbitXLFree on values it returns — that bit is a
        // marker the XLL sets on its own UDF return values — so gating the
        // free on xlbitXLFree (the previous implementation) meant this
        // destructor never freed anything and every Excel-allocated payload
        // (xlGetName / xlfGetCell strings, refs, arrays) leaked. Free the
        // pointer-bearing types unconditionally; scalar types own no
        // Excel-side memory and need no call.
        switch (m_op.xltype & ~(xlbitXLFree | xlbitDLLFree)) {
        case xltypeStr:
        case xltypeMulti:
        case xltypeRef:
            Excel12(xlFree, 0, 1, &m_op);
            break;
        }
        m_op.xltype = xltypeNil; // Prevent double free
    }

    XLOPER12 m_op;
};
