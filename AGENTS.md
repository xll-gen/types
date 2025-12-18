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

## Agent Guidelines

*   **Modification**: When modifying `protocol.fbs`, you must regenerate the Go code and the C++ header (`protocol_generated.h`).
*   **Dependencies**: This library should remain independent of the IPC mechanism (`shm`) and the main `xll-gen` CLI logic.
*   **Formatting**: Follow standard Go and C++ formatting guidelines.
*   **Testing**: Ensure changes are verified. Since this is a library, unit tests (if added) should cover type conversions.
