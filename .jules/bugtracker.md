# Bug Tracker

## 1. Critical Memory Leak in `GridToXLOPER12`

*   **Status:** Resolved
*   **Severity:** Critical
*   **Description:**
    In `GridToXLOPER12` (src/converters.cpp), when converting a FlatBuffers `Grid` to `XLOPER12`, string elements are allocated using `new XCHAR[]`. However, the resulting `XLOPER12` element does not have `xlbitDLLFree` set. While `xlAutoFree12` (src/mem.cpp) correctly iterates through the array and deletes strings if `val.str` is present, there is a risk that if `xlAutoFree12` is not called (e.g., if the root `XLOPER` doesn't have `xlbitDLLFree` set properly, or if Excel manages the memory differently than expected for sub-elements), leaks will occur.

    More critically, `xlAutoFree12` implementation in `src/mem.cpp` assumes ownership of *all* strings inside an `xltypeMulti` array if the array itself is being freed. This is generally correct for this library's usage, but if `GridToXLOPER12` fails halfway through construction (e.g. `bad_alloc`), there is no RAII cleanup, leading to leaks of partially allocated strings.

*   **My Judgment:**
    Implemented a `ScopeGuard` in `src/converters.cpp` that safely cleans up partially initialized arrays and the `XLOPER12` structure itself upon exception. The array is zero-initialized via `memset` to ensure safe cleanup.

## 2. Integer Overflow in Grid Size Calculation

*   **Status:** Resolved
*   **Severity:** High
*   **Description:**
    In `src/converters.cpp`, the calculation `size_t count = (size_t)rows * cols` handles the multiplication safely by promoting to `size_t` (assuming 64-bit). However, `rows` and `cols` are inputs from FlatBuffers. If the FlatBuffer contains extremely large values such that `rows * cols` exceeds `INT_MAX`, the code returns an empty grid or error.

    The issue is potentially in `ConvertGrid` where `rows` and `cols` come from `XLOPER12` (int).
    `size_t count = (size_t)rows * cols;`

    The check `count > (size_t)std::numeric_limits<int>::max()` is present.
    However, the reproduction test `tests/repro_bug.cpp` suggests a specific overflow case might be mishandled or the behavior is unexpected by the user (returning Error vs Empty).

    Current behavior:
    ```cpp
    if (rows < 0 || cols < 0 || count > (size_t)std::numeric_limits<int>::max()) {
         return ... Error ...
    }
    ```
    This seems correct for `GridToXLOPER12`.

    For `ConvertNumGrid` (Excel -> FlatBuffers), if overflow occurs, it returns an empty grid. The user might expect an error or specific handling.

*   **My Judgment:**
    Verified that overflow checks are robust. Added additional `SIZE_MAX` checks for allocation sizes (`SIZE_MAX / sizeof(XLOPER12)`).

## 3. Potential Null Pointer Dereference in `AnyToXLOPER12`

*   **Status:** Resolved
*   **Severity:** Medium
*   **Description:**
    In `AnyToXLOPER12`, `any` is checked for null. However, inside `case protocol::AnyValue::NumGrid`, it calls `ng->data()`.

    ```cpp
    if (... || !ng->data() || ng->data()->size() < count)
    ```
    This check safeguards against null data.

    However, in `ConvertNumGrid` (FP12 -> FBS), `fp` is checked.

    One potential issue is `GridToXLOPER12`.
    ```cpp
    if (!grid) return NULL;
    // ...
    if (... || !grid->data() || ...)
    ```
    This seems safe.

    The potential issue lies in `xlAutoFree12` in `src/mem.cpp`.
    ```cpp
    if (p->val.array.lparray) {
        // ...
        delete[] p->val.array.lparray;
    }
    ```
    It iterates `count` times. `count` is derived from `rows * columns`. If `lparray` was allocated but smaller than `rows*columns` (due to some corruption or bug elsewhere), this reads out of bounds. Trusting `rows*columns` matches `lparray` size is a strong assumption.

*   **My Judgment:**
    The code now strictly ensures allocation matches `rows * cols`. The `ScopeGuard` ensures that if allocation fails or we throw before filling, we clean up exactly what was allocated.

## 4. Missing Exception Handling for `new`

*   **Status:** Resolved
*   **Severity:** High
*   **Description:**
    The code uses raw `new` and `new[]` extensively. If memory allocation fails, `std::bad_alloc` is thrown. Since these functions are likely called from `xlAutoOpen` or as UDFs from Excel, an uncaught C++ exception crashing the DLL is bad behavior (might crash Excel).

    Locations: `src/converters.cpp` (`new XLOPER12[]`, `new XCHAR[]`), `src/mem.cpp` (`new wchar_t[]`), `src/ObjectPool.h` (`new T()`).

*   **My Judgment:**
    Wrapped `AnyToXLOPER12` and `GridToXLOPER12` main bodies in `try-catch (...)` blocks. They now return `xltypeErr` upon any C++ exception (including `bad_alloc`).

## 5. Integer Overflow in `AnyToXLOPER12` (NumGrid)

*   **Status:** Resolved
*   **Severity:** High
*   **Description:**
    In `AnyToXLOPER12` (src/converters.cpp), the `NumGrid` case calculates `count = rows * cols` and then allocates `new XLOPER12[count]`.
    While `count` is checked against `INT_MAX`, there was no check that `count * sizeof(XLOPER12)` does not overflow `size_t`.
    On 32-bit systems (where `size_t` is 32-bit), `count` up to `2*10^9` is allowed, but `count * 32` would wrap around, causing a small allocation and subsequent heap overflow.

*   **My Judgment:**
    Added a check `if (count > SIZE_MAX / sizeof(XLOPER12))` to prevent this overflow.

## 6. String Allocation Denial of Service

*   **Status:** Resolved
*   **Severity:** Medium
*   **Description:**
    In `GridToXLOPER12` (src/converters.cpp), strings are converted using `MultiByteToWideChar` with the full UTF-8 length.
    If the input string is huge (e.g. 1GB), `needed` (wide char count) will be huge.
    The code allocates `new XCHAR[needed + 2]`. This allows a potential attacker (or bad data) to exhaust memory (DoS) or potentially cause integer overflow in `needed + 2`.
    Since Excel cells only support ~32k characters, allocating GBs of memory is wasteful and dangerous.

*   **My Judgment:**
    Added a strict limit check on `needed` (10 million characters or `SIZE_MAX` overflow) before allocation. If the string is too large, it is treated as an empty string (or truncated) to prevent DoS.

## 7. Unsafe NULL Return Values

*   **Status:** Resolved
*   **Severity:** Medium
*   **Description:**
    `GridToXLOPER12`, `RangeToXLOPER12`, and `NumGridToFP12` return `NULL` (nullptr) when inputs are invalid (e.g. `!grid`).
    Callers expecting a valid `XLOPER12` structure might crash if they don't check for NULL.
    `AnyToXLOPER12` returns `GridToXLOPER12(...)` directly. If that returns NULL, `AnyToXLOPER12` returns NULL, potentially crashing consumers.

*   **My Judgment:**
    Modified `GridToXLOPER12` and `RangeToXLOPER12` to return a valid `XLOPER12` with `xltypeErr` (xlerrValue) instead of `NULL`. Modified `NumGridToFP12` to return a valid empty `FP12` (0x0) instead of `NULL`.
