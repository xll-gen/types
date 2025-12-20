## 2024-05-22 - FlatBuffers to Excel String Conversion
**Learning:** `std::wstring_convert` and intermediate allocations in `StringToWString` + `WStringToPascalString` are expensive for large grids of strings.
**Action:** Use `MultiByteToWideChar` to write directly to the target `XLOPER12` buffer, estimating size first. This avoids 2-3 allocations per string.

## 2024-05-23 - NumGrid Bulk Copy
**Learning:** Copying FlatBuffers scalar vectors element-by-element using `Get(i)` is significantly slower (~12x for doubles) than using `std::memcpy`, as `memcpy` leverages optimized assembly for bulk memory transfer.
**Action:** For homogenous numeric data (NumGrid), always verify if the source vector is contiguous (it is in FlatBuffers) and use `memcpy`.

## 2024-05-23 - [FlatBuffers Go Vector Construction]
**Learning:** FlatBuffers Go `StartVector` aligns the buffer but does not reserve space for elements in a way that allows direct `copy` into `Bytes`. `Prep` must be called to move the head, but doing so manually can conflict with `StartVector`'s alignment logic if not careful. Hoisting vtable lookups (`Table().Offset()`) out of the loop is a safer and significant optimization for large vectors.
**Action:** When optimizing FlatBuffers in Go, prefer hoisting lookups over manual buffer manipulation unless you deeply understand the `Builder` state.
