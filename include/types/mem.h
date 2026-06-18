#pragma once
#include <windows.h>
#include "types/xlcall.h"
#include <string>

// Signature for Excel-callback functions exported by the XLL.
// Excel resolves these symbols by literal name from the PE export table,
// so they MUST be declspec(dllexport) even when this translation unit is
// compiled into a static library that is later linked into the XLL.
#define TYPES_EXCEL_CALLBACK extern "C" __declspec(dllexport) void __stdcall

/**
 * Allocates an XLOPER12 from the thread-safe object pool and initializes it to empty.
 *
 * @return Pointer to a pool-managed XLOPER12, freed by xlAutoFree12.
 */
LPXLOPER12 NewXLOPER12();

/**
 * Releases an XLOPER12 back to the pool without freeing its content.
 *
 * Internal use only (e.g. for async handlers that extract values).
 *
 * @param p XLOPER12 previously obtained from NewXLOPER12().
 */
void ReleaseXLOPER12(LPXLOPER12 p);

/**
 * Creates an XLOPER12 String (Pascal-style wide string) managed by the DLL.
 *
 * Sets xltypeStr | xlbitDLLFree. The returned pointer and the string buffer
 * are both managed and will be freed by xlAutoFree12.
 *
 * @param str Source string to copy into Excel-managed storage.
 * @return Pointer to a managed XLOPER12 string.
 */
LPXLOPER12 NewExcelString(const std::wstring& str);

/**
 * Creates an FP12 array managed by a ring buffer (valid for return to Excel).
 *
 * @note FP12 is used with the "K%" type. Excel copies the data, so it only
 *       needs to persist until return.
 *
 * @param rows Number of rows.
 * @param cols Number of columns.
 * @return Pointer to a ring-buffer-managed FP12 array.
 */
FP12* NewFP12(int rows, int cols);

/**
 * Frees ONLY the DLL-owned heap buffers hanging off an XLOPER12's value union
 * — the `xltypeStr` buffer, the `xltypeMulti` element strings plus the element
 * array, or the `xltypeRef` mref buffer. It does NOT release the XLOPER12 struct
 * itself (the pool / ReleaseXLOPER12 is the caller's job) and never touches
 * borrowed / Excel-owned pointers. Within a multi, an element string is freed
 * only when it carries `xlbitDLLFree` (our ownership marker).
 *
 * This is the single definition of the ownership-critical free logic that used
 * to be copy-pasted between `xlAutoFree12` (mem.cpp) and the `GridToXLOPER12`
 * cleanup guard (converters.cpp) — a pair that had to stay byte-identical or
 * one side leaks while the other double-frees (BUG-014/015/017 lineage).
 *
 * PRECONDITION (multi case): the element array must be zero-initialized or fully
 * constructed. Do NOT call this on a multi whose elements still have
 * indeterminate `xltype` (e.g. straight after `new XLOPER12[n]` with no memset)
 * — the per-element `xlbitDLLFree` test would read garbage. Excel hands
 * `xlAutoFree12` complete XLOPERs; `GridToXLOPER12` memsets its array before any
 * path that can reach its guard. Build sites that allocate a numeric grid but do
 * NOT zero it keep their own array-only `delete[]` (see converters.cpp).
 */
void FreeDllOwnedContents(LPXLOPER12 p);

/**
 * Callback called by Excel to free memory allocated by the XLL.
 *
 * Must be exported by literal name (`xlAutoFree12`) from the final XLL.
 *
 * @param p XLOPER12 whose DLL-managed memory Excel is returning for release.
 */
TYPES_EXCEL_CALLBACK xlAutoFree12(LPXLOPER12 p);
