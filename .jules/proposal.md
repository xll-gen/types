# Proposals for xll-gen/types

## 1. Enable Compiler Warnings
**Status:** Completed
**Description:** The `CMakeLists.txt` was updated to enable `-Wall -Wextra -Wpedantic` (or `/W4` on MSVC). This exposed unused parameter warnings in `src/win_compat.cpp`, which were subsequently fixed using `[[maybe_unused]]`.
**Result:** Build is now cleaner and stricter.

## 2. Replace Deprecated `std::wstring_convert`
**Status:** Proposed
**Description:** `src/win_compat.cpp` uses `std::wstring_convert` and `std::codecvt_utf8`, which are deprecated in C++17.
**Impact:** Future compiler versions may remove these, breaking the Linux build (which relies on `win_compat`).
**Plan:** Refactor `MultiByteToWideChar` and `WideCharToMultiByte` implementation in `win_compat.cpp` to use other mechanisms (e.g., `mbstowcs` with UTF-8 locale, or a lightweight library).

## 3. Implement `GetXllDir` for Linux
**Status:** Proposed
**Description:** `src/utility.cpp` implements `GetXllDir` using `GetModuleFileNameW`. On Linux (via `win_compat`), this returns 0/empty string.
**Impact:** Tests relying on file paths relative to the library location may fail or require workarounds on Linux.
**Plan:** Implement `GetXllDir` using `/proc/self/exe` (Linux) or `dladdr` to find the shared library or executable path.

## 4. Optimize Go DeepCopy for Byte Vectors
**Status:** Proposed
**Description:** In `go/protocol/deepcopy.go`, `AsyncHandle` (and potentially other future byte vectors) are copied byte-by-byte in a loop.
**Impact:** Inefficient for large data.
**Plan:** Use `CreateByteVector` (if available in generated bindings) or `CreateByteVector` logic to copy slices directly.

## 5. Fix `XlError` Mapping
**Status:** Completed
**Description:** `XlError` enum values in `protocol.fbs` start at 2000, while Excel error codes start at 0. The C++ conversion logic was casting them directly, causing invalid error codes to be sent to/from Excel.
**Action:** Added `ProtocolErrorToExcel` and `ExcelErrorToProtocol` helpers in `src/converters.cpp` to correctly offset the values by 2000. Updated `tests/test_converters.cpp` to verify correct mapping using real Excel values.

## 6. Support `RefCache` and `AsyncHandle` in `AnyToXLOPER12`
**Status:** Proposed
**Description:** `AnyToXLOPER12` currently does not handle `RefCache` and `AsyncHandle` types, defaulting to `xltypeNil`.
**Impact:** Potential data loss or incorrect behavior if these types are returned to Excel.
**Plan:** Determine appropriate mapping (e.g., `xltypeBigData` or string representation) and implement.

## 7. Safety Hardening for C++ Converters
**Status:** In Progress
**Description:** `ConvertGrid`, `ConvertNumGrid`, and `ConvertScalar` lack internal exception handling. If memory allocation fails (e.g., `std::bad_alloc` during `vector::reserve` or `builder::CreateString`), the exception propagates, potentially crashing the host application (Excel).
**Plan:** Wrap these function bodies in `try-catch` blocks and return a safe "empty" or "error" FlatBuffer offset on failure.

## 8. Safety Hardening for Go DeepCopy
**Status:** In Progress
**Description:** Generated `DeepCopy` methods (e.g., for `Grid`, `NumGrid`) allocate memory based on the input `DataLength` field without validation. A malicious or malformed payload with a large length field could cause an Out-Of-Memory (OOM) panic (DoS).
**Plan:** Add validation checks in `DeepCopy` methods to ensure `DataLength` is within reasonable limits (e.g., `math.MaxInt32`) before allocation.

## 9. Parallel execution for Go validation tests
**Status:** Completed
**Date:** 2025-12-23
**Description:** `TestRangeValidation` and `TestGridOverflowValidation` in `go/protocol/validation_test.go` were missing `t.Parallel()`, preventing them from running concurrently with other tests.
**Action:** Added `t.Parallel()` to these tests.
**Result:** Tests now run in parallel.

## 10. Code cleanup in GridToXLOPER12
**Status:** Completed
**Date:** 2025-12-23
**Description:** `GridToXLOPER12` in `src/converters.cpp` used a manual cleanup loop in its `ScopeGuard` that duplicated logic found in `xlAutoFree12`.
**Action:** Replaced the manual loop with a call to `xlAutoFree12(op)`.
**Result:** Reduced code duplication and improved maintainability.

## 11. Consolidate String Conversion Logic
**Status:** Proposed
**Date:** 2025-12-23
**Description:** `GridToXLOPER12` implements custom `MultiByteToWideChar` logic with truncation and buffer size management. `src/utility.cpp` provides `StringToWString` but it doesn't support the specific Excel truncation (32767 chars) logic used in the converter.
**Plan:** Create a specialized utility function (e.g., `Utf8ToExcelString`) in `src/utility.cpp` that encapsulates the truncation and optimization logic from `converters.cpp`. This would clean up `converters.cpp` and make the logic reusable.
