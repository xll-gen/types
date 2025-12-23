# xll-gen/types

This repository contains the core types and conversion logic for `xll-gen`. It provides:
1.  **Go Protocol Types**: Generated FlatBuffer Go code for the `xll-gen` protocol.
2.  **C++ Library**: A static library for converting between Excel types (`XLOPER12`, `FP12`) and FlatBuffers, along with memory management utilities.
3.  **FlatBuffers Schema**: The `protocol.fbs` schema definition.

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

## Co-Change Clusters

Certain parts of the codebase are tightly coupled and must be updated together to preserve consistency.

### Protocol Schema & Generated Code
The `go/protocol/protocol.fbs` definition is the single source of truth for the data structures.
1.  **Schema Source**: `go/protocol/protocol.fbs`.
2.  **Generated C++**: `include/types/protocol_generated.h` must be regenerated via `flatc`.
3.  **Generated Go**: `go/protocol/*.go` files must be regenerated via `flatc`.
**Constraint**: Any change to `go/protocol/protocol.fbs` requires running the generation task (`task generate` or similar) to update both C++ and Go artifacts in the same commit.

### Excel Conversion Logic
When adding support for a new type in `protocol.fbs`:
1.  **Schema**: Update `go/protocol/protocol.fbs`.
2.  **C++ Converters**: Update `src/converters.cpp` and `include/types/converters.h` to map the new FlatBuffer type to `XLOPER12`.
3.  **Go Helpers**: Update `go/protocol/extensions.go` or validation logic if the new type requires specific handling on the Go side.
4.  **Deep Copy**: Update `go/protocol/deepcopy.go` if the new type requires manual deep copy logic (e.g., it contains pointers or nested structures).

### Memory Management
1.  **Allocator**: `src/mem.cpp` (implementing `xlAutoFree12`).
2.  **Object Pool**: `include/types/ObjectPool.h`.
**Constraint**: Changes to how `XLOPER12`s are allocated or freed must be reflected in both the allocation strategy (ObjectPool) and the cleanup callback (`xlAutoFree12`).

## Agent Guidelines

*   **Modification**: When modifying `protocol.fbs`, you must regenerate the Go code and the C++ header (`protocol_generated.h`).
*   **Dependencies**: This library should remain independent of the IPC mechanism (`shm`) and the main `xll-gen` CLI logic.
*   **Formatting**: Follow standard Go and C++ formatting guidelines.
*   **Testing**: Ensure changes are verified. Since this is a library, unit tests (if added) should cover type conversions.
*   **Workflow**: Before committing, fetch the latest state of the parent branch and merge it into your working branch.
*   **Documentation**: Before writing to documents in the `.jules` directory, check the current date and ensure it is included in the entry.
