# FetchContent for FlatBuffers
include(FetchContent)

FetchContent_Declare(
    flatbuffers
    GIT_REPOSITORY https://github.com/google/flatbuffers.git
    GIT_TAG        v25.9.23
)

FetchContent_MakeAvailable(flatbuffers)
