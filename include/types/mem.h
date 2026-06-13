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
 * Callback called by Excel to free memory allocated by the XLL.
 *
 * Must be exported by literal name (`xlAutoFree12`) from the final XLL.
 *
 * @param p XLOPER12 whose DLL-managed memory Excel is returning for release.
 */
TYPES_EXCEL_CALLBACK xlAutoFree12(LPXLOPER12 p);
