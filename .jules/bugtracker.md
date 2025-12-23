# Bug Tracker

## 14. Memory Leak in `RangeToXLOPER12`

*   **Status:** Open
*   **Severity:** Low
*   **Description:**
    In `src/converters.cpp`, `RangeToXLOPER12` allocates `op->val.mref.lpmref` using `new`. If an exception occurs (e.g. `std::bad_alloc` later), the `ScopeGuard` releases `op` but fails to free the allocated `lpmref` buffer, causing a leak.
*   **My Judgment:**
    This is a definite leak under error conditions. While low severity (requires OOM to trigger), it violates RAII principles. It should be fixed by ensuring the `ScopeGuard` or a separate mechanism frees `lpmref` if `op` is released.

## 15. Integer Overflow in `WideToUtf8`

*   **Status:** Open
*   **Severity:** Low
*   **Description:**
    In `src/utility.cpp`, `WideToUtf8` casts `wstring::size` to `int`. For strings > 2GB (unlikely but possible), this overflows, causing invalid arguments to `WideCharToMultiByte`.
*   **My Judgment:**
    This is a robustness issue. While 2GB strings are rare in Excel, the library should fail safely (throw or return empty) rather than invoking UB or crashing due to overflow. A simple size check is recommended.

## 16. Unsafe API Exposure (`ConvertGrid`)

*   **Status:** Open
*   **Severity:** Medium
*   **Description:**
    `ConvertGrid` in `src/converters.cpp` allocates memory based on input dimensions without a `try-catch` block. Malicious inputs can trigger `std::bad_alloc` and crash the host application.
*   **My Judgment:**
    Public APIs expected to be called by host applications should catch internal exceptions to prevent crashing the process. This function should be wrapped in `try-catch` and return a safe error value.

## 17. Memory Leak in `AnyToXLOPER12` (NumGrid)

*   **Status:** Open
*   **Severity:** Medium
*   **Description:**
    In `src/converters.cpp` (NumGrid case), `op` is acquired, then `lparray` is allocated via `new`. If `new` throws, `ScopeGuard` (declared after) is not active, leaking `op`.
*   **My Judgment:**
    This is a clear RAII ordering bug. `ScopeGuard` must be declared immediately after the resource (`op`) is acquired to ensure it is released if subsequent operations (like `new`) fail.

## 18. DoS and Panic in Go DeepCopy

*   **Status:** Open
*   **Severity:** High
*   **Description:**
    In `go/protocol/deepcopy.go`, `DeepCopy` iterates using `DataLength()` without bounds checking against the actual buffer size. Malformed FlatBuffers can cause:
    1. Runtime Panic (Index out of range) -> Service Crash.
    2. Memory Exhaustion (DoS) due to huge `Make` calls based on `Length`.
*   **My Judgment:**
    This is a high-severity security vulnerability. A remote attacker could crash the Go service or exhaust its memory by sending a small packet with a large `Length` field. Validation logic must be added to `DeepCopy` or the accessor methods to ensure `Length` is backed by actual data and within safe limits.
