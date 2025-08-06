cmake_minimum_required(VERSION 3.10)
project(XLSXColumnExtractor)

set(CMAKE_CXX_STANDARD 17)

if (MSVC)
  message(STATUS "Forçando runtime /MT (linkagem estática)")
  foreach(flag_var 
    CMAKE_C_FLAGS 
    CMAKE_C_FLAGS_DEBUG 
    CMAKE_C_FLAGS_RELEASE 
    CMAKE_C_FLAGS_MINSIZEREL 
    CMAKE_C_FLAGS_RELWITHDEBINFO 
    CMAKE_CXX_FLAGS 
    CMAKE_CXX_FLAGS_DEBUG 
    CMAKE_CXX_FLAGS_RELEASE 
    CMAKE_CXX_FLAGS_MINSIZEREL 
    CMAKE_CXX_FLAGS_RELWITHDEBINFO)
    
    string(REPLACE "/MD" "/MT" ${flag_var} "${${flag_var}}")
  endforeach()
endif()

add_subdirectory(libs/OpenXLSX)

add_executable(xlsx-column main.cpp)
target_link_libraries(xlsx-column OpenXLSX)

# Tentar link estático no GCC/MinGW
if (CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    set_target_properties(xlsx-column PROPERTIES
        LINK_FLAGS "-static -static-libgcc -static-libstdc++"
    )
endif()

# Se for MSVC (Visual Studio), use runtime estático
if (MSVC)
    foreach(flag_var
            CMAKE_C_FLAGS_RELEASE
            CMAKE_C_FLAGS_DEBUG
            CMAKE_CXX_FLAGS_RELEASE
            CMAKE_CXX_FLAGS_DEBUG)
        string(REPLACE "/MD" "/MT" ${flag_var} "${${flag_var}}")
    endforeach()
endif()
