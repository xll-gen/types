## 2025-12-23 - FlatBuffers to Excel String Conversion
**Learning:** `std::wstring_convert` and intermediate allocations in `StringToWString` + `WStringToPascalString` are expensive for large grids of strings.
**Action:** Use `MultiByteToWideChar` to write directly to the target `XLOPER12` buffer, estimating size first. This avoids 2-3 allocations per string.

## 2025-12-23 - NumGrid Bulk Copy
**Learning:** Copying FlatBuffers scalar vectors element-by-element using `Get(i)` is significantly slower (~12x for doubles) than using `std::memcpy`, as `memcpy` leverages optimized assembly for bulk memory transfer.
**Action:** For homogenous numeric data (NumGrid), always verify if the source vector is contiguous (it is in FlatBuffers) and use `memcpy`.

## 2025-12-23 - [FlatBuffers Go NumGrid Optimization Rejected]
**Learning:** Hoisting vtable lookups in `NumGrid.DeepCopy` offered ~30% speedup but required hardcoding the field offset (8). This creates a fragile dependency on the generated schema that is too risky for maintenance.
**Action:** Do not implement optimizations that rely on hardcoded generated field offsets.

## 2025-12-23 - Excel String Conversion Optimization
**Learning:** Intermediate `std::wstring` allocations in string conversion helpers (like `ConvertExcelString`) add significant overhead (malloc + copy) for high-frequency operations.
**Action:** Use platform APIs (like `WideCharToMultiByte`) directly on the source buffer when possible to eliminate intermediate objects.

## 2025-12-23 - FlatBuffers Vector Creation
**Learning:** `builder.CreateVector(std::vector)` incurs double allocation (std::vector + FlatBuffer) and double copy. Using `CreateUninitializedVector` allows writing directly to the FlatBuffer memory.
**Action:** When populating FlatBuffer vectors from non-contiguous sources (where `memcpy` isn't possible), use `CreateUninitializedVector` and fill the buffer in a loop to save one allocation and one copy pass.

## 2025-12-23 - [Optimization] GridToXLOPER12 Small String Conversion
**Learning:** `MultiByteToWideChar` is typically called twice (size query + conversion). For small strings (common in Excel cells), a stack buffer can serve as the destination for the first attempt, often eliminating the second call entirely.
**Action:** Use a "speculative" stack buffer (e.g. 256 chars) for `MultiByteToWideChar`. If it succeeds, copy to the final heap buffer. If not, fallback to the two-pass approach.
