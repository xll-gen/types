# FetchContent for FlatBuffers
include(FetchContent)

if(WIN32)
    # Define the FlatBuffers version tag from FetchContent_Declare
    set(FB_VER_TAG "v25.9.23")

    # 1. Try to find installed flatc in PATH
    find_program(FLATC_EXECUTABLE flatc)

    # 2. If not found, try to use a cached downloaded binary
    if(NOT FLATC_EXECUTABLE)
        set(FLATC_DIR "${CMAKE_CURRENT_BINARY_DIR}/flatc_bin/${FB_VER_TAG}") # Version-specific cache directory
        set(CACHED_FLATC_PATH "${FLATC_DIR}/flatc.exe")

        if(EXISTS "${CACHED_FLATC_PATH}")
            set(FLATC_EXECUTABLE "${CACHED_FLATC_PATH}" CACHE FILEPATH "Path to flatc executable" FORCE)
            message(STATUS "Found cached flatc: ${FLATC_EXECUTABLE}")
        else()
            # 3. If not cached, download pre-built binary
            set(FLATC_ZIP "${CMAKE_CURRENT_BINARY_DIR}/flatc_${FB_VER_TAG}_win.zip")

            message(STATUS "flatc not found in PATH or cache. Attempting to download flatc binary from GitHub (${FB_VER_TAG})...")
            file(DOWNLOAD
                "https://github.com/google/flatbuffers/releases/download/${FB_VER_TAG}/Windows.flatc.binary.zip"
                "${FLATC_ZIP}"
                STATUS DOWNLOAD_STATUS
                TIMEOUT 60 # Set a timeout for download
            )
            list(GET DOWNLOAD_STATUS 0 STATUS_CODE)
            list(GET DOWNLOAD_STATUS 1 ERROR_MESSAGE)

            if(STATUS_CODE EQUAL 0)
                file(MAKE_DIRECTORY "${FLATC_DIR}")
                file(ARCHIVE_EXTRACT INPUT "${FLATC_ZIP}" DESTINATION "${FLATC_DIR}")
                file(REMOVE "${FLATC_ZIP}") # Clean up the zip file

                if(EXISTS "${CACHED_FLATC_PATH}")
                    set(FLATC_EXECUTABLE "${CACHED_FLATC_PATH}" CACHE FILEPATH "Path to flatc executable" FORCE)
                    message(STATUS "Successfully downloaded and extracted flatc: ${FLATC_EXECUTABLE}")
                else()
                    message(WARNING "Downloaded zip but 'flatc.exe' not found inside '${FLATC_DIR}'. Will build from source.")
                endif()
            else()
                message(WARNING "Failed to download flatc binary for ${FB_VER_TAG}. Status: ${STATUS_CODE}, Message: ${ERROR_MESSAGE}. Will build from source.")
            endif()
        endif()
    endif()

    if(FLATC_EXECUTABLE)
        # Ensure flatc is not built from source if we have an executable
        set(FLATBUFFERS_BUILD_FLATC OFF CACHE BOOL "Don't build flatc compiler" FORCE)
        # Point FlatBuffers build system to our flatc executable
        set(FLATBUFFERS_FLATC_EXECUTABLE ${FLATC_EXECUTABLE} CACHE FILEPATH "Path to flatc" FORCE)

        # Create an imported target for flatc if it doesn't exist
        # This allows other CMake targets to depend on it or use $<TARGET_FILE:flatc>
        if(NOT TARGET flatc)
            add_executable(flatc IMPORTED GLOBAL)
            set_target_properties(flatc PROPERTIES IMPORTED_LOCATION "${FLATC_EXECUTABLE}")
        endif()
    else()
        message(STATUS "flatc executable not found or downloaded. FlatBuffers will build flatc from source.")
    endif()
endif()

FetchContent_Declare(
    flatbuffers
    GIT_REPOSITORY https://github.com/google/flatbuffers.git
    GIT_TAG        ${FB_VER_TAG} # Use the version tag defined above
)

set(FLATBUFFERS_BUILD_TESTS OFF CACHE BOOL "Disable FlatBuffers tests" FORCE)

FetchContent_MakeAvailable(flatbuffers)
