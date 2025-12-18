## 2024-05-22 - FlatBuffers to Excel String Conversion
**Learning:** `std::wstring_convert` and intermediate allocations in `StringToWString` + `WStringToPascalString` are expensive for large grids of strings.
**Action:** Use `MultiByteToWideChar` to write directly to the target `XLOPER12` buffer, estimating size first. This avoids 2-3 allocations per string.
