# xll-gen Types & Conversion Library

This repository contains the core types and conversion logic for the `xll-gen` project. It serves two main purposes:
1.  **Go Protocol Types**: Provides generated FlatBuffers Go code for the `xll-gen` protocol.
2.  **C++ Library**: Provides a static library for converting between Excel types (`XLOPER12`, `FP12`) and FlatBuffers, along with memory management utilities.

## Contents

- [Go Protocol Types](#go-protocol-types)
- [C++ Library](#c-library)
  - [Installation](#installation)
  - [API Reference](#api-reference)
    - [Type Conversions](#type-conversions)
    - [Memory Management](#memory-management)
    - [String Utilities](#string-utilities)
    - [General Utilities](#general-utilities)
    - [Object Pool](#object-pool)
    - [Excel SDK](#excel-sdk)

## Go Protocol Types

The `go/protocol` directory contains the generated Go code from the `protocol.fbs` FlatBuffers schema (located at `go/protocol/protocol.fbs`). This package is used to serialize and deserialize messages exchanged between the Excel add-in and the backend service.

### Usage

```go
import "github.com/xll-gen/types/go/protocol"
```

The package provides structs and builders for all protocol messages (e.g., `Range`, `Scalar`, `Grid`, `Command`).

## C++ Library

The C++ library provides the heavy lifting for marshalling data between Excel's native formats and FlatBuffers. It also handles memory management for Excel data structures.

### Installation

You can include this library in your CMake project using `FetchContent`.

```cmake
include(FetchContent)
FetchContent_Declare(
    xll-gen-types
    GIT_REPOSITORY https://github.com/xll-gen/types.git
    GIT_TAG        main # Replace with specific tag or commit
)
FetchContent_MakeAvailable(xll-gen-types)

target_link_libraries(your_target PRIVATE xll-gen-types)
```

**Dependencies**:
- **FlatBuffers**: This project depends on Google FlatBuffers. The CMake configuration handles this dependency.

### Cross-Platform Support

To facilitate unit testing and development on non-Windows platforms (e.g., Linux, macOS), this library includes a compatibility layer (`include/types/win_compat.h`). This layer mocks necessary Windows data types (like `DWORD`, `HANDLE`) and Excel structures (`XLOPER12`) so that the library can be compiled and tested without the Windows SDK or Excel installed.

The compatibility layer is automatically included when compiling on non-Windows systems.

### API Reference

#### Type Conversions

Header: `include/types/converters.h`

These functions convert between Excel `XLOPER12`/`FP12` structures and FlatBuffers objects.

**Excel to FlatBuffers:**

*   `flatbuffers::Offset<protocol::Range> ConvertRange(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder, const std::string& format = "")`
    *   Converts an `XLOPER12` (which can be a single value, array, or reference) into a `protocol::Range` message.
*   `flatbuffers::Offset<protocol::Scalar> ConvertScalar(const XLOPER12& cell, flatbuffers::FlatBufferBuilder& builder)`
    *   Converts a single `XLOPER12` cell value to a `protocol::Scalar`.
*   `flatbuffers::Offset<protocol::Grid> ConvertGrid(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder)`
    *   Converts an `XLOPER12` array/multi to a `protocol::Grid` (2D array of scalars).
*   `flatbuffers::Offset<protocol::NumGrid> ConvertNumGrid(FP12* fp, flatbuffers::FlatBufferBuilder& builder)`
    *   Converts an `FP12` (floating point array) to a `protocol::NumGrid`.
*   `flatbuffers::Offset<protocol::Any> ConvertAny(LPXLOPER12 op, flatbuffers::FlatBufferBuilder& builder)`
    *   Generic conversion that detects the type of `XLOPER12` and converts it to the appropriate `protocol::Any` union type.

**FlatBuffers to Excel:**

*   `LPXLOPER12 AnyToXLOPER12(const protocol::Any* any)`
    *   Converts a `protocol::Any` message back to an `XLOPER12`. The returned pointer is managed by the library's memory pool or is an Excel-managed string.
*   `LPXLOPER12 RangeToXLOPER12(const protocol::Range* range)`
    *   Converts a `protocol::Range` to `XLOPER12`.
*   `LPXLOPER12 GridToXLOPER12(const protocol::Grid* grid)`
    *   Converts a `protocol::Grid` to `XLOPER12`.
*   `FP12* NumGridToFP12(const protocol::NumGrid* grid)`
    *   Converts a `protocol::NumGrid` to `FP12`.

#### Memory Management

Header: `include/types/mem.h`

Manages memory for `XLOPER12` and `FP12` structures, ensuring thread safety and proper cleanup.

*   `LPXLOPER12 NewXLOPER12()`
    *   Allocates a new `XLOPER12` from the thread-safe object pool.
*   `void ReleaseXLOPER12(LPXLOPER12 p)`
    *   Returns an `XLOPER12` to the pool. Use this for internal cleanup of intermediate values.
*   `LPXLOPER12 NewExcelString(const std::wstring& str)`
    *   Creates an `XLOPER12` string (Pascal-style) that is managed by the DLL. It sets `xlbitDLLFree`, so Excel will call `xlAutoFree12` when done.
*   `FP12* NewFP12(int rows, int cols)`
    *   Allocates a new `FP12` structure.
*   `void __stdcall xlAutoFree12(LPXLOPER12 p)`
    *   The standard callback invoked by Excel to free memory allocated by the XLL (specifically for `xlbitDLLFree` types).

#### String Utilities

Header: `include/types/PascalString.h`

Utilities for handling Excel's length-prefixed (Pascal) strings.

*   `std::vector<char> CStringToPascalString(const std::string& c_str)`
    *   Converts C++ `std::string` to a byte-length-prefixed string.
*   `std::string PascalStringToCString(const unsigned short* pascal_str)`
    *   Converts a byte-length-prefixed string to `std::string`.
*   `std::vector<wchar_t> WStringToPascalString(const std::wstring& w_str)`
    *   Converts `std::wstring` to Excel 12's wide character Pascal string (first `wchar_t` is length).
*   `std::wstring PascalString12ToWString(const wchar_t* pascal_str)`
    *   Converts Excel 12 Pascal string to `std::wstring`.
*   `std::wstring PascalToWString(const wchar_t* pascal_str)`
    *   Alias for `PascalString12ToWString`.
*   `wchar_t* WStringToNewPascalString(const std::wstring& w_str)`
    *   Creates a new Pascal string on the heap.

#### General Utilities

Header: `include/types/utility.h`

*   `LPXLOPER12 TempStr12(const wchar_t* txt)`
    *   Creates a temporary `XLOPER12` string (mostly for calling Excel functions).
*   `LPXLOPER12 TempInt12(int val)`
    *   Creates a temporary `XLOPER12` integer.
*   `std::wstring StringToWString(const std::string& str)`
*   `std::string WideToUtf8(const std::wstring& wstr)`
*   `std::wstring ConvertToWString(const char* str)`
*   `std::string ConvertExcelString(const wchar_t* wstr)`
*   `bool IsSingleCell(LPXLOPER12 pxRef)`
    *   Checks if a reference `XLOPER12` points to a single cell.
*   `std::wstring GetXllDir()`
    *   Returns the directory where the XLL is located.
*   `void SetDebugFlag(bool enabled)`
    *   Enables or disables debug logging.
*   `bool GetDebugFlag()`
    *   Returns the current debug logging state.
*   `void DebugLog(const char* fmt, ...)`
    *   Logs a formatted message to the debug output (Visual Studio Output window on Windows, stderr on other platforms) if the debug flag is enabled.

#### Object Pool

Header: `include/types/ObjectPool.h`

*   `template <typename T, size_t ShardCount = 16> class ObjectPool`
    *   A thread-safe, sharded object pool used internally for `XLOPER12` allocation to reduce heap contention.

#### Excel SDK

Header: `include/types/xlcall.h`

This library includes the standard Excel C API definitions (SDK Version 15.0), defining types like `XLOPER12`, `FP12`, `XLREF12`, etc., and the `Excel12` / `Excel12v` callbacks.

## Development

This project uses [Task](https://taskfile.dev) to automate development tasks.

### Prerequisites

*   CMake 3.24+
*   Go 1.18+
*   C++ Compiler (C++17 support)
*   `clang-format` (optional, for formatting)
*   [Task](https://taskfile.dev/installation/)

### Commands

*   **Build**: `task build` (configures and builds).
*   **Test**: `task test` (runs unit tests).
*   **Format**: `task format` (formats C++ and Go files).
*   **Generate**: `task generate` (regenerates Go and C++ code).
*   **Clean**: `task clean` (removes build directory).

If you don't have `task` installed, you can run the underlying CMake commands directly (see `Taskfile.yml` for details).
