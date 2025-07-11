cmake_minimum_required(VERSION 3.10)

# Test configuration
enable_testing()

# Include directories for tests
include_directories(${CMAKE_CURRENT_SOURCE_DIR}/../include)
include_directories(${CMAKE_CURRENT_SOURCE_DIR}/../third_party)

# Test sources
set(TEST_SOURCES
    test_syntax_checker.cpp
    test_cli_parser.cpp
    test_macro_runner.cpp
)

# Create test executable for each test file
foreach(TEST_SOURCE ${TEST_SOURCES})
    get_filename_component(TEST_NAME ${TEST_SOURCE} NAME_WE)
    add_executable(${TEST_NAME} ${TEST_SOURCE})
    
    # Link against the main project libraries
    target_link_libraries(${TEST_NAME}
        -lm
        -ldl
        -pthread
    )
    
    # Add the test
    add_test(NAME ${TEST_NAME} COMMAND ${TEST_NAME})
endforeach()

# Integration tests
add_executable(integration_tests integration_tests.cpp)
target_link_libraries(integration_tests
    -lm
    -ldl
    -pthread
)
add_test(NAME integration_tests COMMAND integration_tests)