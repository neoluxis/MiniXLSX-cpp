cmake_minimum_required(VERSION 3.11)

project(MiniXLSX)

option(MiniXLSX_BUILD_TEST ON)
option(MiniXLSX_BUILD_DEMO ON)


set(CMAKE_CXX_STANDARD 20)

set(MLSX_SRCS
    src/XLDocument.cpp
)





add_library(MiniXLSX SHARED ${MLSX_SRCS})


add_subdirectory(KFZippa)

if(MiniXLSX_BUILD_TEST)
    add_subdirectory(tests)
endif(MiniXLSX_BUILD_TEST)

if(MiniXLSX_BUILD_DEMO)
    add_subdirectory(demos)
endif(MiniXLSX_BUILD_DEMO)