# Changelog

All notable changes to `xll-gen/types` are documented here. (File introduced
at v0.2.9 — for earlier releases see the git tag history.)

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
