## 2024-05-22 - FlatBuffers to Excel String Conversion
**Learning:** `std::wstring_convert` and intermediate allocations in `StringToWString` + `WStringToPascalString` are expensive for large grids of strings.
**Action:** Use `MultiByteToWideChar` to write directly to the target `XLOPER12` buffer, estimating size first. This avoids 2-3 allocations per string.

## 2024-05-23 - NumGrid Bulk Copy
**Learning:** Copying FlatBuffers scalar vectors element-by-element using `Get(i)` is significantly slower (~12x for doubles) than using `std::memcpy`, as `memcpy` leverages optimized assembly for bulk memory transfer.
**Action:** For homogenous numeric data (NumGrid), always verify if the source vector is contiguous (it is in FlatBuffers) and use `memcpy`.

## 2024-05-23 - [FlatBuffers Go NumGrid Optimization Rejected]
**Learning:** Hoisting vtable lookups in `NumGrid.DeepCopy` offered ~30% speedup but required hardcoding the field offset (8). This creates a fragile dependency on the generated schema that is too risky for maintenance.
**Action:** Do not implement optimizations that rely on hardcoded generated field offsets.
