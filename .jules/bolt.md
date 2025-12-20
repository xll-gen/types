## 2024-05-22 - FlatBuffers to Excel String Conversion
**Learning:** `std::wstring_convert` and intermediate allocations in `StringToWString` + `WStringToPascalString` are expensive for large grids of strings.
**Action:** Use `MultiByteToWideChar` to write directly to the target `XLOPER12` buffer, estimating size first. This avoids 2-3 allocations per string.

## 2024-05-23 - NumGrid Bulk Copy
**Learning:** Copying FlatBuffers scalar vectors element-by-element using `Get(i)` is significantly slower (~12x for doubles) than using `std::memcpy`, as `memcpy` leverages optimized assembly for bulk memory transfer.
**Action:** For homogenous numeric data (NumGrid), always verify if the source vector is contiguous (it is in FlatBuffers) and use `memcpy`.
