# Sentinel Journal

## 2024-05-22 - XLOPER12 Null Pointer Dereference
**Vulnerability:** `ConvertGrid` and `ConvertMultiToAny` in `src/converters.cpp` accessed `op->val.array.lparray` without checking for NULL, assuming that if `rows * cols > 0`, the array pointer is valid. A malformed `XLOPER12` with positive dimensions but a NULL array pointer caused a segmentation fault (DoS).
**Learning:** External C-style structures (like `XLOPER12`) must be treated as untrusted input. Logical consistency (dimensions imply data) is not guaranteed.
**Prevention:** Explicitly validate all pointer fields in `XLOPER12` structures against `nullptr` before dereferencing, regardless of other metadata (like row/column counts).

## 2025-12-23 - FlatBuffers Go DeepCopy DoS
**Vulnerability:** `DeepCopy` methods in Go blindly trusted `DataLength()` from the FlatBuffer and used it to allocate slices via `make`. A malicious or corrupted FlatBuffer with a huge length field caused `make` to panic or OOM, crashing the application.
**Learning:** Go's `make` panics on large sizes (even if they fit in `int` but exceed memory/implementation limits). Also, generated FlatBuffers code accessors trust the underlying data. Manual deep copy logic must validate lengths.
**Prevention:** Always validate vector lengths against a reasonable limit (e.g. `math.MaxInt32`) before allocation in deep copy or deserialization logic.
