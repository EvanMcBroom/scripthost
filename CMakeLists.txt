cmake_minimum_required(VERSION 3.14.0)
project(scripthost CXX)

# Statically build targets
string(REGEX REPLACE "/M[TD]d?" "" CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG}")
set(CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG} /MTd")
string(REGEX REPLACE "/M[TD]d?" "" CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE}")
set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} /MT")

add_executable(scripthost)
set_target_properties(scripthost PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED YES
)
target_sources(scripthost PRIVATE comfactory.cpp comserver.cpp comserver.def main.cpp scriptsite.cpp)
# Manually add def file to link options because CMake ignores def files for executables
target_link_options(scripthost PRIVATE "/DEF:${CMAKE_CURRENT_SOURCE_DIR}/comserver.def")
add_subdirectory(include)