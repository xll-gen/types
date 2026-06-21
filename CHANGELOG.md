# Changelog

All notable changes to `xll-gen/types` are documented here. (File introduced
at v0.2.9 — for earlier releases see the git tag history.)

## [v0.2.14] - 2026-06-22

### Changed

- **Pascal wide-string encoding unified (R33).** The XLOPER12 `xltypeStr` layout
  ("clamp to 32767, `[0]`=len, copy body, NUL-terminate") that was hand-rolled in
  five sites (`ScopedXLOPER12::SetString`, the `const XLOPER12*` ctor,
  `NewExcelString`, and `Utf8ToExcelString`'s stack + heap paths) now flows
  through one header-only writer, `WritePascalWString` (+ `WritePascalWBufferLen`
  for sizing, sharing a single `ClampExcelStringLen`). One intentional
  normalization: the `const XLOPER12*` ctor used to copy a source length prefix
  verbatim, so a malformed prefix claiming e.g. 40000 produced a buffer
  advertising 40000 over a 32767-char body; it now stamps the clamped length.
  Well-formed input is byte-identical. XLOPER12 ABI unchanged.
- Internal converter/clone refactors (R28–R32): unified XLOPER12 free-contents
  logic, `Clone()` boilerplate collapsed via a generic `cloneTable`, grid
  alloc-overflow guard + scalar XLOPER12 makers consolidated, and union
  completeness guards added for `ScalarValue`/`AnyValue`.

### Fixed

- `deepcopy`: fail closed on an inaccessible `RtdConnectRequest.Strings` element
  instead of silently proceeding.

## [v0.2.12] - 2026-06-13

### Changed

- **The C++ library is now Windows-only.** All non-Windows (Linux/macOS/POSIX)
  portability support has been removed. Every header now `#include <windows.h>`
  unconditionally; the `#ifdef _WIN32 ... #else #include "types/win_compat.h"
  ... #endif` pattern is gone from `converters.h`, `mem.h`, `ScopedXLOPER12.h`,
  and `utility.h`, and the matching `#ifdef`/`#else` branches in `utility.cpp`
  (`GetXllDir`, `DebugLog`, the `__linux__` includes) and `xlcall.cpp` collapse
  to the Windows path. Tests drop their POSIX guards likewise. Supported
  toolchains are MSVC 2019+ and MinGW (the `windows-mingw` CMake preset).

### Fixed

- **`deepcopy.go` deep-copy now fails closed on inaccessible vector elements.**
  `Grid`/`BatchRtdUpdate` return offset 0 (the existing graceful-failure
  convention) instead of emitting a vector with a null (0) offset hole, and
  `Range.DeepCopy` pre-validates every `Rect` before opening the inline-struct
  vector so a mid-build access failure can no longer leave `EndVector` with a
  short vector and corrupt the buffer. No public signature change; behavior
  differs only on corrupt/truncated input. New regression `TestRange_Clone_MultiRef`.

### Removed

- **`include/types/win_compat.h` and `src/win_compat.cpp`** — the test-only
  shim that faked `<windows.h>` for non-Windows compile validation. The
  `XLLGEN_TYPES_TESTING` define and the `if(NOT WIN32)` branch that propagated
  it in `CMakeLists.txt` are also removed.

### Docs

- `mem.h` API declarations converted to Doxygen `/** */` blocks with
  `@param`/`@return`/`@note` tags.

## [v0.2.10] - 2026-06-12

### Added

- **`ConvertRange` now populates `Range.sheet_name`** (`"[Book]Sheet"`) via
  the thread-safe-callable `xlSheetNm`/`xlSheetId` C-API entry points — legal
  from non-macro, `$`-registered worksheet functions (required by xll-gen
  v0.5.0's caller/macro registration split, where `caller: true` no longer
  registers `#`). Resolved once per `ConvertRange` call; degrades gracefully
  to an empty name outside a live Excel calc context (e.g. unit tests).
  Purely additive — the field was previously always empty. Excel-allocated
  lookup results are released via `ScopedXLOPER12Result` (the v0.2.9
  contract). New graceful-degradation tests in `tests/test_roundtrip.cpp`.

## [v0.2.9] - 2026-06-12

### Fixed

- **`ScopedXLOPER12Result` destructor never freed Excel-allocated results.**
  `Free()` was gated on `xlbitXLFree`, but Excel never sets that bit on values
  it returns from `Excel12()` — it is a marker the XLL sets on its own UDF
  return values. The destructor therefore never released anything, leaking
  every Excel-allocated payload held by the wrapper (`xlGetName` strings,
  `xlfGetCell` format strings, refs, arrays) across xll-gen generated code.
  `Free()` now releases pointer-bearing types (`xltypeStr`/`xltypeMulti`/
  `xltypeRef`) unconditionally via `xlFree` — the documented caller-side
  contract for `Excel12` results — and skips scalars, which own no Excel-side
  memory. `tests/test_scoped_xloper.cpp` updated to pin the new contract,
  including the previously-untested "no marker bit, must still free" case.
  Documented invariant: only `Excel12()` result operands may be stored in a
  `Result` (DLL-allocated values belong to `xlAutoFree12`).

## [v0.2.8] and earlier

See `git log` / GitHub release notes.
