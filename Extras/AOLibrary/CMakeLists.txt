cmake_minimum_required(VERSION 3.10)

# set the project name
project(AOLibrary VERSION 1.0)

if (NOT MSVC)
	message(FATAL_ERROR "¡Necesitas un compilador de Visual Studio!")
endif()

# specify the C++ standard
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED True)
set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS True)

# set compiler flags
set(CMAKE_CXX_FLAGS "/W4")

# Get a list of all *.cpp and files to pass to the compiler
file(GLOB SOURCE_CODE CONFIGURE_DEPENDS "src/*.cpp")

# Declare library and sources
add_library(${PROJECT_NAME} SHARED ${SOURCE_CODE})

# https://blog.kitware.com/create-dlls-on-windows-without-declspec-using-new-cmake-export-all-feature/
include(GenerateExportHeader)
generate_export_header(${PROJECT_NAME})

# Include header files from `include` directory
target_include_directories(${PROJECT_NAME} PRIVATE "${CMAKE_SOURCE_DIR}/src/include")