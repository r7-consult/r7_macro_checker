cmake_minimum_required(VERSION 3.10)
project(onlyoffice-macro-cli VERSION 2.0.0)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Build configuration
option(BUILD_TESTS "Build test programs" OFF)
option(USE_V8 "Enable V8 JavaScript engine" ON)
option(USE_JSC "Enable JavaScriptCore engine" ON)

# Find packages
find_package(PkgConfig REQUIRED)

# V8 support
if(USE_V8)
    # Try to find V8 in common locations
    find_path(V8_INCLUDE_DIR 
        NAMES v8.h
        PATHS 
            /usr/include/v8 
            /usr/local/include/v8
            /opt/homebrew/include/v8
            ${CMAKE_SOURCE_DIR}/../core/DesktopEditor/doctrenderer/js_internal/v8
        DOC "V8 include directory"
    )
    
    find_library(V8_LIBRARY 
        NAMES v8 libv8
        PATHS 
            /usr/lib 
            /usr/local/lib
            /opt/homebrew/lib
        DOC "V8 library"
    )
    
    find_library(V8_PLATFORM_LIBRARY 
        NAMES v8_libplatform libv8_libplatform
        PATHS 
            /usr/lib 
            /usr/local/lib
            /opt/homebrew/lib
        DOC "V8 platform library"
    )
    
    if(V8_INCLUDE_DIR AND V8_LIBRARY AND V8_PLATFORM_LIBRARY)
        message(STATUS "Found V8: ${V8_LIBRARY}")
        add_definitions(-DHAVE_V8)
        set(V8_LIBRARIES ${V8_LIBRARY} ${V8_PLATFORM_LIBRARY})
        include_directories(${V8_INCLUDE_DIR})
        set(V8_FOUND TRUE)
    else()
        message(WARNING "V8 not found, V8 engine will be disabled")
        set(V8_FOUND FALSE)
    endif()
endif()

# JavaScriptCore support (primarily for Apple platforms)
if(USE_JSC)
    if(APPLE)
        find_library(JSC_LIBRARY JavaScriptCore)
        if(JSC_LIBRARY)
            message(STATUS "Found JavaScriptCore: ${JSC_LIBRARY}")
            add_definitions(-DHAVE_JSC)
            set(JSC_LIBRARIES ${JSC_LIBRARY})
            set(JSC_FOUND TRUE)
        else()
            message(WARNING "JavaScriptCore not found on Apple platform")
            set(JSC_FOUND FALSE)
        endif()
    else()
        # Try to find JavaScriptCore on non-Apple platforms
        find_path(JSC_INCLUDE_DIR 
            NAMES JavaScriptCore/JavaScriptCore.h
            PATHS /usr/include /usr/local/include
        )
        
        find_library(JSC_LIBRARY 
            NAMES javascriptcoregtk-4.0 
            PATHS /usr/lib /usr/local/lib
        )
        
        if(JSC_INCLUDE_DIR AND JSC_LIBRARY)
            message(STATUS "Found JavaScriptCore GTK: ${JSC_LIBRARY}")
            add_definitions(-DHAVE_JSC)
            set(JSC_LIBRARIES ${JSC_LIBRARY})
            include_directories(${JSC_INCLUDE_DIR})
            set(JSC_FOUND TRUE)
        else()
            message(WARNING "JavaScriptCore not found, JSC engine will be disabled")
            set(JSC_FOUND FALSE)
        endif()
    endif()
endif()

# Check if at least one engine is available
if(NOT V8_FOUND AND NOT JSC_FOUND)
    message(FATAL_ERROR "No JavaScript engine found. Please install V8 or JavaScriptCore development packages.")
endif()

# Source files
set(COMMON_SOURCES
    src/main.cpp
    src/cli_parser.cpp
    src/linter.cpp
    src/linter_utils.cpp
    src/api_definitions.cpp
    src/function_rules.cpp
    src/mistake_rules.cpp
    src/onlyoffice_rules.cpp
    src/style_rules.cpp
    src/variable_rules.cpp
    src/syntax_rules.cpp
    src/syntax_checker.cpp
    src/macro_runner.cpp
    src/js_engine_factory.cpp
)

# Engine-specific sources
set(ENGINE_SOURCES)

if(V8_FOUND)
    list(APPEND ENGINE_SOURCES src/v8_engine.cpp)
endif()

if(JSC_FOUND)
    list(APPEND ENGINE_SOURCES src/jsc_engine.cpp)
endif()

# Remove old Duktape files (they should not be compiled)
# src/duktape_engine.cpp - removed from build
# src/syntax_rules.cpp - replaced with engine-based validation

# All sources
set(ALL_SOURCES ${COMMON_SOURCES} ${ENGINE_SOURCES})

# Create executable
add_executable(onlyoffice-macro-cli ${ALL_SOURCES})

# Include directories
target_include_directories(onlyoffice-macro-cli PRIVATE 
    include
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)

# Link libraries
target_link_libraries(onlyoffice-macro-cli)

if(V8_FOUND)
    target_link_libraries(onlyoffice-macro-cli ${V8_LIBRARIES})
    target_include_directories(onlyoffice-macro-cli PRIVATE ${V8_INCLUDE_DIR})
endif()

if(JSC_FOUND)
    target_link_libraries(onlyoffice-macro-cli ${JSC_LIBRARIES})
    if(JSC_INCLUDE_DIR)
        target_include_directories(onlyoffice-macro-cli PRIVATE ${JSC_INCLUDE_DIR})
    endif()
endif()

# Platform-specific linking
if(UNIX AND NOT APPLE)
    target_link_libraries(onlyoffice-macro-cli pthread dl)
endif()

# Compiler-specific flags
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    target_compile_options(onlyoffice-macro-cli PRIVATE 
        -Wall -Wextra -Wpedantic
        -Wno-unused-parameter
        -fno-rtti  # V8 doesn't use RTTI
    )
endif()

if(MSVC)
    target_compile_options(onlyoffice-macro-cli PRIVATE /W4)
    target_compile_definitions(onlyoffice-macro-cli PRIVATE 
        _CRT_SECURE_NO_WARNINGS
        NOMINMAX
    )
endif()

# Build configuration information
message(STATUS "Build configuration:")
message(STATUS "  V8 Engine: ${V8_FOUND}")
message(STATUS "  JavaScriptCore Engine: ${JSC_FOUND}")
message(STATUS "  Build Tests: ${BUILD_TESTS}")

# Tests (if enabled)
if(BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()

# Install
install(TARGETS onlyoffice-macro-cli 
    RUNTIME DESTINATION bin
)

# Install examples
install(DIRECTORY examples/ 
    DESTINATION share/onlyoffice-macro-cli/examples
    FILES_MATCHING PATTERN "*.js"
)

# Install documentation
install(FILES README.md 
    DESTINATION share/doc/onlyoffice-macro-cli
)

# Package configuration
set(CPACK_PACKAGE_NAME "onlyoffice-macro-cli")
set(CPACK_PACKAGE_VERSION ${PROJECT_VERSION})
set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "OnlyOffice JavaScript Macro CLI Tool")
set(CPACK_PACKAGE_VENDOR "ONLYOFFICE")
set(CPACK_PACKAGE_CONTACT "support@onlyoffice.com")

if(UNIX)
    set(CPACK_GENERATOR "DEB;RPM;TGZ")
    set(CPACK_DEBIAN_PACKAGE_DEPENDS "libc6")
    set(CPACK_RPM_PACKAGE_REQUIRES "glibc")
endif()

if(WIN32)
    set(CPACK_GENERATOR "NSIS;ZIP")
endif()

include(CPack)