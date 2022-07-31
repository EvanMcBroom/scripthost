cmake_minimum_required(VERSION 3.14.0)
project(scripthost CXX)

# Statically build executable targets
string(REGEX REPLACE "/M[TD]d?" "" CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG}")
set(CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG} /MTd")
string(REGEX REPLACE "/M[TD]d?" "" CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE}")
set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} /MT")

# Create an object library from the sources for the scripting site implementation
# They will be linked twice and the object library will reduce build time
add_library(scriptingsite OBJECT comfactory.cpp comserver.cpp scriptsite.cpp)
set_target_properties(scriptingsite PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED YES
    POSITION_INDEPENDENT_CODE YES # Build as PIC to support being linked into shared libraries
)

# Build the com server for the scripting site implementation
add_library(server SHARED comserver.def $<TARGET_OBJECTS:scriptingsite>)

# Build the standalone executable with the scripting site embedded
add_executable(scripthost)
set_target_properties(scripthost PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED YES
)
target_sources(scripthost PRIVATE main.cpp $<TARGET_OBJECTS:scriptingsite>)
target_link_libraries(scripthost PRIVATE scriptingsite)

add_subdirectory(include)