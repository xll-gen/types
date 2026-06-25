# xll-gen/types

This repository contains the core types and conversion logic for `xll-gen`. It provides:
1.  **Go Protocol Types**: Generated FlatBuffer Go code for the `xll-gen` protocol.
2.  **C++ Library**: A static library for converting between Excel types (`XLOPER12`, `FP12`) and FlatBuffers, along with memory management utilities.
3.  **FlatBuffers Schema**: The `protocol.fbs` schema definition.

## Companion Repos

This module is one of four coordinated repos. When work crosses boundaries, read the relevant `AGENTS.md`:

* **`github.com/xll-gen/xll-gen`** — XLL generator; pins this module to an exact version in both Go (`internal/generator/dependencies.go`) and C++ (`CMakeLists.txt.tmpl` `GIT_TAG`). Bumps here require a coordinated bump there.
* **`github.com/xll-gen/shm`** — transport. Strictly independent of this module: `shm` must not import `types`.
* **`github.com/xll-gen/sugar`** — unrelated; not in the runtime path.

## Platform Targets

* **Go code (`go/protocol/`)**: portable; runs on any GOOS/GOARCH supported by Go and the FlatBuffers generated code.
* **C++ code (`include/types/`, `src/`)**: **Windows-only.** Targets Windows x86 / x86-64 (Intel/AMD) with the Excel SDK, built with MSVC 2019+ or MinGW (the `windows-mingw` CMake preset). The headers `#include <windows.h>` unconditionally; there is no non-Windows / Linux / macOS portability layer. The former `win_compat.h` shim has been **removed** — do not reintroduce it.
* **No ARM**: the production deployment chain (`xll-gen` runtime) is Windows x86/x64 only.

## Structure

*   `go/protocol`: Go package containing the generated FlatBuffers code.
*   `include/types`: C++ header files.
*   `src`: C++ source files.
*   `cmake`: CMake modules.

## Usage

### Go

```go
import "github.com/xll-gen/types/go/protocol"
```

### C++

Add this repository as a dependency in your CMake project (e.g., using `FetchContent`).

```cmake
include(FetchContent)
FetchContent_Declare(
    xll-gen-types
    GIT_REPOSITORY https://github.com/xll-gen/types.git
    GIT_TAG        main
)
FetchContent_MakeAvailable(xll-gen-types)

target_link_libraries(your_target PRIVATE xll-gen-types)
```

## Development

This project uses [Task](https://taskfile.dev/) for build automation.

### Prerequisites

1.  **Task**: Download the `task` binary and place it in the root directory (or install it globally).
    *   Do NOT commit the `task` binary.
2.  **FlatBuffers Compiler (`flatc`)**: Download `flatc` and place it in the root directory (or ensure it is in your PATH).
    *   Do NOT commit the `flatc` binary.

### Commands

*   **Build**: `./task build`
*   **Test**: `./task test`
*   **Clean**: `./task clean`
*   **Format**: `./task fmt`
*   **All**: `./task` (runs configure, build, and test)

## Design Decisions & Limitations

*   **CommandWrapper**: The `CommandWrapper` table is used in `protocol.fbs` because FlatBuffers Go bindings do not currently support vectors of Unions (e.g., `[Command]`). The wrapper allows us to use `[CommandWrapper]` instead.

*   **`xltype` bit flags are the source of truth for "is this guard redundant?" (over-defensive-logic audit, 2026-06-25).** `types` (via `xlcall.h`) pins `xltypeRef = 0x0008` and `xltypeSRef = 0x0400` as **distinct, non-overlapping bits**. A 2026-06-25 cross-repo audit confirmed a downstream consumer guard (`CacheManager::GetOrComputeRefHash` in `xll-gen`) is **not** redundant precisely because of this: its caller admits `xltypeRef | xltypeSRef`, but the callee narrows to `xltypeRef`, so an SRef-only input passes the caller yet is correctly filtered by the callee — real filtering, not a duplicate check. **General rule for any consumer of these types:** a callee guard is redundant only when the caller's admission predicate *logically implies* it. With distinct bit flags, a narrower callee mask is load-bearing — do not "simplify" it away by reasoning only from the call graph.

## Documentation Standards

*   **Timestamps**: When creating or updating documentation in the `.jules/` directory (e.g., bug tracker, sentinel logs), always fetch the current date (e.g., via `date` command) and use it to timestamp your entries. This ensures accurate tracking of when issues or changes were recorded.

## Co-Change Clusters

Certain parts of the codebase are tightly coupled and must be updated together to preserve consistency.

### Protocol Schema & Generated Code
The `go/protocol/protocol.fbs` definition is the single source of truth for the data structures.
1.  **Schema Source**: `go/protocol/protocol.fbs`.
2.  **Generated C++**: `include/types/protocol_generated.h` must be regenerated via `flatc`.
3.  **Generated Go**: `go/protocol/*.go` files must be regenerated via `flatc`.
**Constraint**: Any change to `go/protocol/protocol.fbs` requires running the generation task (`task generate` or similar) to update both C++ and Go artifacts in the same commit.

**Confirmed-correct note — the `Date` union member is NOT consumed by `xll-gen` generated projects (do not "remove as dead", do not assume it gates XLL date support; verified 2026-06-16):** `xll-gen` maps its `date` type to `SchemaType="double"`, so a generated XLL encodes dates as plain `double` and never references `protocol::Date`. The generated project's protocol code comes from this `types` module (pinned by version), and `xll-gen`'s template `protocol.fbs` is only a flatc parse-stub. Therefore the `Date{serial,format}` table here is currently unused by the generated-XLL path; its presence/absence in `xll-gen`'s template does not affect generated-project compilation. Keep it (it is the schema's source of truth and may back a future non-double encoding), but understand it is not on the live date-I/O path today.

### Excel Conversion Logic
When adding support for a new type in `protocol.fbs`:
> **Completeness guards (R29)**: appending a member to the `ScalarValue`/`AnyValue` unions trips a **loud failure** — the C++ `static_assert(…::MAX == …::Date)` before `AnyToXLOPER12`/`GridToXLOPER12`'s switches fails to compile, and `TestUnionDeepCopyCompleteness` (Go) fails on the changed member count. Both error messages list every union-tag ladder to update (`AnyToXLOPER12`, `GridToXLOPER12` per-cell, `ConvertScalar`/`ConvertAny`, and `deepcopy.go`'s two switches). Update all of them, then bump the asserts/count. These were added because the ladders fall through to Nil on an unknown tag — a missed ladder used to degrade silently. Do not delete the guards to "make it build"; update the ladders.
1.  **Schema**: Update `go/protocol/protocol.fbs`.
2.  **C++ Converters**: Update `src/converters.cpp` and `include/types/converters.h` to map the new FlatBuffer type to `XLOPER12`.
3.  **Go Helpers**: Update `go/protocol/extensions.go` or validation logic if the new type requires specific handling on the Go side.
4.  **Deep Copy**: Update `go/protocol/deepcopy.go` if the new type requires manual deep copy logic (e.g., it contains pointers or nested structures). The `Clone()` skeleton is now generic: each `Clone()` is one line — `if rcv == nil { return nil }; return cloneTable(rcv, GetRootAsX)` (see R30 below). A new root table only needs that one-liner plus a `DeepCopy` body; do **not** re-roll the round-trip-through-a-builder boilerplate.

### Memory Management
1.  **Allocator**: `src/mem.cpp` (implementing `xlAutoFree12`).
2.  **Object Pool**: `include/types/ObjectPool.h`.
**Constraint**: Changes to how `XLOPER12`s are allocated or freed must be reflected in both the allocation strategy (ObjectPool) and the cleanup callback (`xlAutoFree12`).
**Single free-contents helper (R28)**: The DLL-owned-contents teardown (str buffer / multi element-strings + array / ref mref) lives once in `FreeDllOwnedContents(LPXLOPER12)` (`include/types/mem.h` / `src/mem.cpp`); `xlAutoFree12` and the `GridToXLOPER12` cleanup guard both call it (the pool release stays the caller's job). When adding a new XLOPER12 shape that owns heap memory, extend `FreeDllOwnedContents` — do not open-code a parallel free path. The element-string test uses the base-type mask `(xltype & ~(xlbitDLLFree|xlbitXLFree)) == xltypeStr` (not a loose `& xltypeStr`) to avoid type-confused frees. One deliberate carve-out: the `NumGridToFP12` guard keeps an array-only `delete[]` because it fills uninitialized elements (see the in-code note).

## Agent Guidelines

*   **Modification**: When modifying `protocol.fbs`, you must regenerate the Go code and the C++ header (`protocol_generated.h`).
*   **Dependencies**: This library should remain independent of the IPC mechanism (`shm`) and the main `xll-gen` CLI logic.
*   **Formatting**: Follow standard Go and C++ formatting guidelines.
*   **Testing**: Ensure changes are verified. Since this is a library, unit tests (if added) should cover type conversions.
*   **Workflow**: Before committing, fetch the latest state of the parent branch and merge it into your working branch.
*   **Documentation**: Before writing to documents in the `.jules` directory, check the current date and ensure it is included in the entry.

## Known Improvement Backlog

From a code review on 2026-05-16. Address as part of normal work.

* **Builder helpers**: the generated FlatBuffers code exposes `*Builder` types, but there are no hand-written helpers in `go/protocol/extensions.go` that hide the `flatbuffers.Builder` ceremony. Add `BuildScalar(b, v)`, `BuildGrid(b, rows, cols, data)` etc. for the common shapes — `xll-gen/pkg/server` is currently re-implementing this dance and is the canonical caller to migrate.
* **DONE (v0.2.5):** Deep copy performance — `go/protocol/builder_pool.go` adds a `sync.Pool` of `*flatbuffers.Builder`. All 10 `Clone()` methods in `deepcopy.go` now `acquireBuilder()` + `defer releaseBuilder(b)` instead of `flatbuffers.NewBuilder(0)`. Sustained workloads converge to one builder per goroutine in the pool; under xll-gen's async batcher and RTD handler hot paths this was the dominant alloc per Clone.
* **DONE (post-0.2.13 / R30):** `Clone()` boilerplate collapsed via generic `cloneTable[T any](rcv deepCopier, getRoot func([]byte, flatbuffers.UOffsetT) *T) *T` (`deepcopy.go`). All 10 `Clone()` methods are now `if rcv == nil { return nil }; return cloneTable(rcv, GetRootAsX)` (nil guard kept at call site so a typed-nil never enters `DeepCopy`). Net −47 lines. Scope: skeleton only — the per-vector length / fail-closed security guards inside each `DeepCopy` body were untouched.
* **DONE (post-0.2.13 / R28):** XLOPER12 free-contents unified into `FreeDllOwnedContents` (see Memory Management constraint above). Two byte-identical cleanup copies (`xlAutoFree12`, `GridToXLOPER12` guard) collapsed to one; element-string test hardened to a base-type mask.
* **DONE (post-0.2.13):** `RtdConnectRequest.DeepCopy` Strings vector now fails closed (`return 0`) on an inaccessible/nil element, matching the other vectors (`Grid.Data`, `Range.Refs`, `BatchRtdUpdate.Updates`) — previously a truncated buffer silently emitted an empty string. (A legitimate empty string is a non-nil zero-length slice, so valid data is unaffected.)
* **DONE (v0.2.5):** Regenerate provenance — `go/protocol/flatc_version.go` carries `const FlatcVersion = "25.9.23"`. `flatc_version_test.go::TestFlatcVersion_GeneratedFilesPresent` walks `go/protocol/*.go` and verifies each carries the FlatBuffers `Code generated` banner (catches a partial regeneration). xll-gen's `cmd/doctor` (or its `TestFlatbuffersVersionConsistency` test) cross-checks against `internal/versions.FlatBuffers`.
* **SUPERSEDED:** C++ `win_compat.h` (the non-Windows compile shim) has been **removed entirely** — the repo is now Windows-only (MSVC 2019+ / MinGW). All headers `#include <windows.h>` unconditionally; there is no `XLLGEN_TYPES_TESTING` define and no Linux/macOS build path. Do not reintroduce a portability shim.
* **DONE (v0.2.6):** C++ round-trip coverage — `tests/test_roundtrip.cpp` (wired up as the `roundtrip_test` CTest target) builds an `XLOPER12 → FlatBuffer → XLOPER12` fixture per converter (Scalar, Grid, NumGrid, Range, Any) and diffs the result. 35 test cases cover NaN/±Inf/±0.0 doubles, `INT_MIN`/`INT_MAX`, UTF-16 non-ASCII strings, empty/1×1/2×2 grids, and every `AnyValue` union variant. Documents one protocol limitation surfaced by the suite: `table Num { val: double; }`'s default-value elision drops the sign on -0.0 (the `[double]` vector in NumGrid is unaffected).

## CLAUDE.md / Agent Tool Compatibility

This repository is configured so that AI tools using `CLAUDE.md` (Claude Code) read this `AGENTS.md` as the authoritative source. **All durable agent guidance must live here, not in `CLAUDE.md`.** `CLAUDE.md`, if present, must contain only a one-line redirect to this file.
