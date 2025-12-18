cmake_minimum_required(VERSION 3.24)

file(READ "VERSION" VER_RAW)
string(STRIP "${VER_RAW}" VER)

project(xll-gen-types VERSION ${VER} LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")
include(flatbuffers)

# Define the library
add_library(xll-gen-types STATIC
    src/converters.cpp
    src/mem.cpp
    src/PascalString.cpp
    src/utility.cpp
    src/win_compat.cpp
    src/xlcall.cpp
)

# Include directories
target_include_directories(xll-gen-types PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:include>
)

# Link dependencies
target_link_libraries(xll-gen-types PUBLIC flatbuffers)

# Tests
enable_testing()
add_subdirectory(tests)

# Install rules (optional, if you want to install)
include(GNUInstallDirs)
install(TARGETS xll-gen-types
    EXPORT xll-gen-types-targets
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)
install(DIRECTORY include/types DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

# Custom Targets

# 1. Generate FlatBuffers code
# We rely on 'flatc' target being available from flatbuffers dependency
if(TARGET flatc)
    add_custom_target(generate
        COMMAND $<TARGET_FILE:flatc> --go -o ${CMAKE_CURRENT_SOURCE_DIR}/go/ ${CMAKE_CURRENT_SOURCE_DIR}/go/protocol/protocol.fbs
        COMMAND $<TARGET_FILE:flatc> --cpp --scoped-enums -o ${CMAKE_CURRENT_SOURCE_DIR}/include/types ${CMAKE_CURRENT_SOURCE_DIR}/go/protocol/protocol.fbs
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Regenerating FlatBuffers code (Go and C++)"
    )
endif()

# 2. Format code
find_program(CLANG_FORMAT clang-format)
find_program(GO_EXEC go)

add_custom_target(format)

if(CLANG_FORMAT)
    # Glob sources to format
    file(GLOB_RECURSE CPP_SOURCES
        "${CMAKE_CURRENT_SOURCE_DIR}/src/*.cpp"
        "${CMAKE_CURRENT_SOURCE_DIR}/src/*.h"
        "${CMAKE_CURRENT_SOURCE_DIR}/include/*.h"
        "${CMAKE_CURRENT_SOURCE_DIR}/tests/*.cpp"
    )
    add_custom_target(format_cpp
        COMMAND ${CLANG_FORMAT} -i ${CPP_SOURCES}
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Formatting C++ sources"
    )
    add_dependencies(format format_cpp)
endif()

if(GO_EXEC)
    add_custom_target(format_go
        COMMAND ${GO_EXEC} fmt ./...
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        COMMENT "Formatting Go sources"
    )
    add_dependencies(format format_go)
endif()
