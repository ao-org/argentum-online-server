cmake_minimum_required(VERSION 3.20.2)

# Set the project name
project(AOLibrary VERSION 1.0)

if (NOT MSVC)
	message(FATAL_ERROR "¡Necesitas un compilador de Visual Studio!")
endif()

# Specify the C++ standard
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED True)

# Tell CMake to create export file for all symbols
set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS True)

# Set compiler flags
string(APPEND CMAKE_CXX_FLAGS "/W4")

# Get a list of all *.cpp and files to pass to the compiler
file(GLOB SOURCE_CODE CONFIGURE_DEPENDS "src/*.cpp")

# Declare library and sources
add_library(${PROJECT_NAME} SHARED ${SOURCE_CODE})

# https://blog.kitware.com/create-dlls-on-windows-without-declspec-using-new-cmake-export-all-feature/
include(GenerateExportHeader)
generate_export_header(${PROJECT_NAME})

# Include header files from `include` directory
target_include_directories(${PROJECT_NAME} PRIVATE "${CMAKE_SOURCE_DIR}/src/include")