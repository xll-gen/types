cmake_minimum_required(VERSION 3.24)
project(xll-gen-types VERSION 0.1.0 LANGUAGES CXX)

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
    src/xlcall.cpp
)

# Include directories
target_include_directories(xll-gen-types PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:include>
)

# Link dependencies
target_link_libraries(xll-gen-types PUBLIC flatbuffers::flatbuffers)

# Install rules (optional, if you want to install)
include(GNUInstallDirs)
install(TARGETS xll-gen-types
    EXPORT xll-gen-types-targets
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)
install(DIRECTORY include/types DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})
