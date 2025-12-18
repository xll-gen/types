add_executable(types_test test_converters.cpp)
target_link_libraries(types_test PRIVATE xll-gen-types)
target_include_directories(types_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME types_test COMMAND types_test)

add_executable(repro_test repro_bug.cpp)
target_link_libraries(repro_test PRIVATE xll-gen-types)
target_include_directories(repro_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME repro_test COMMAND repro_test)

add_executable(logging_test test_logging.cpp)
target_link_libraries(logging_test PRIVATE xll-gen-types)
target_include_directories(logging_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME logging_test COMMAND logging_test)

add_executable(security_test test_security_regressions.cpp)
target_link_libraries(security_test PRIVATE xll-gen-types)
target_include_directories(security_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME security_test COMMAND security_test)
