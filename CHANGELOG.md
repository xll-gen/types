# Changelog

All notable changes to `xll-gen/types` are documented here. (File introduced
at v0.2.9 — for earlier releases see the git tag history.)

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
