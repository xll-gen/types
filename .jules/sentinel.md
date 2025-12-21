# Sentinel Journal

## 2024-05-22 - XLOPER12 Null Pointer Dereference
**Vulnerability:** `ConvertGrid` and `ConvertMultiToAny` in `src/converters.cpp` accessed `op->val.array.lparray` without checking for NULL, assuming that if `rows * cols > 0`, the array pointer is valid. A malformed `XLOPER12` with positive dimensions but a NULL array pointer caused a segmentation fault (DoS).
**Learning:** External C-style structures (like `XLOPER12`) must be treated as untrusted input. Logical consistency (dimensions imply data) is not guaranteed.
**Prevention:** Explicitly validate all pointer fields in `XLOPER12` structures against `nullptr` before dereferencing, regardless of other metadata (like row/column counts).
