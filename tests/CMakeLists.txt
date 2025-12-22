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

add_executable(fixes_test test_fixes.cpp)
target_link_libraries(fixes_test PRIVATE xll-gen-types)
target_include_directories(fixes_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME fixes_test COMMAND fixes_test)

add_executable(null_safety_test test_null_safety.cpp)
target_link_libraries(null_safety_test PRIVATE xll-gen-types)
target_include_directories(null_safety_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME null_safety_test COMMAND null_safety_test)

add_executable(utility_safety_test test_utility_safety.cpp)
target_link_libraries(utility_safety_test PRIVATE xll-gen-types)
target_include_directories(utility_safety_test PRIVATE ${CMAKE_SOURCE_DIR}/include)
add_test(NAME utility_safety_test COMMAND utility_safety_test)
