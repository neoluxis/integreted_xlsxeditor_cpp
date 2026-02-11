cmake_minimum_required(VERSION 3.11)

project(XLSXEditor)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_BUILD_TYPE Debug)

message("Integrated XLSX editor in femapp")

# Find Qt6
find_package(Qt6 REQUIRED COMPONENTS Core Widgets LinguistTools)

# Enable Qt MOC, RCC, UIC
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOUIC_SEARCH_PATHS ui)

# If not already added (e.g., in FemApp), add MiniXLSX
if(NOT TARGET MiniXLSX)
    set(MiniXLSX_BUILD_TEST OFF CACHE BOOL "Build MiniXLSX test applications" FORCE)
    set(MiniXLSX_BUILD_DEMO OFF CACHE BOOL "Build MiniXLSX demo applications" FORCE)
    set(MiniXLSX_BUILD_TOOL OFF CACHE BOOL "Build MiniXLSX tool applications" FORCE)
    add_subdirectory(3rdparty/MiniXLSX)
endif()

# Wrap UI files
qt6_wrap_ui(UI_HEADERS
    src/DataItem.ui
    src/XLSXEditor.ui
)

# Set sources
set(WIDGET_SRC
    src/XLSXEditor.cpp
    src/DataItem.cpp
    include/XLSXEditor.hpp
    include/DataItem.hpp
    ${UI_HEADERS}
)

# For testing, if main.cpp exists
if(EXISTS ${CMAKE_CURRENT_SOURCE_DIR}/src/main.cpp)
    set(EXE_SRC
        src/main.cpp
    )
endif()

# Create library
add_library(XLSXEditor ${WIDGET_SRC})

# Include directories
target_include_directories(XLSXEditor PUBLIC
    include
    ${CMAKE_CURRENT_BINARY_DIR}  # For generated UI headers
    3rdparty/MiniXLSX/include/cc/neolux/utils/MiniXLSX
)

# Link libraries
target_link_libraries(XLSXEditor PUBLIC
    Qt6::Core
    Qt6::Widgets
    MiniXLSX
)

# Translation files
FILE(GLOB TRANSLATION_TS translations/*.ts)

# Update translations
add_custom_target(update_translations_XLSXEditor
    COMMAND $<TARGET_FILE:Qt6::lupdate>
    ${WIDGET_SRC}
    -ts ${TRANSLATION_TS}
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    COMMENT "Update Qt translation files for XLSXEditor (.ts)"
)

# Compile translations
add_custom_target(compile_translations_XLSXEditor
    COMMAND $<TARGET_FILE:Qt6::lrelease> ${TRANSLATION_TS}
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    COMMENT "Compile Qt translation files for XLSXEditor (.qm)"
)

# If building exe for testing
if(EXE_SRC)
    add_executable(XLSXEditor_test ${EXE_SRC})
    target_link_libraries(XLSXEditor_test PRIVATE XLSXEditor)
endif()

