# Proposals

## [Fixed] Parallel execution for Go validation tests
- **Issue**: `TestRangeValidation` and `TestGridOverflowValidation` in `go/protocol/validation_test.go` were missing `t.Parallel()`, preventing them from running concurrently with other tests.
- **Action**: Added `t.Parallel()` to these tests.
- **Result**: Tests now run in parallel.

## [Fixed] Code cleanup in GridToXLOPER12
- **Issue**: `GridToXLOPER12` in `src/converters.cpp` used a manual cleanup loop in its `ScopeGuard` that duplicated logic found in `xlAutoFree12`.
- **Action**: Replaced the manual loop with a call to `xlAutoFree12(op)`.
- **Result**: Reduced code duplication and improved maintainability.

## [Proposed] Consolidate String Conversion Logic
- **Issue**: `GridToXLOPER12` implements custom `MultiByteToWideChar` logic with truncation and buffer size management. `src/utility.cpp` provides `StringToWString` but it doesn't support the specific Excel truncation (32767 chars) logic used in the converter.
- **Proposal**: Create a specialized utility function (e.g., `Utf8ToExcelString`) in `src/utility.cpp` that encapsulates the truncation and optimization logic from `converters.cpp`. This would clean up `converters.cpp` and make the logic reusable.
