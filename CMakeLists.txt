cmake_minimum_required(VERSION 3.11)

project(MiniXLSX)

option(MiniXLSX_BUILD_TEST "Build MiniXLSX test applications" ON)
option(MiniXLSX_BUILD_DEMO "Build MiniXLSX demo applications" ON)


set(CMAKE_CXX_STANDARD 20)

find_package(PkgConfig REQUIRED)
pkg_check_modules(PUGIXML REQUIRED pugixml)

set(MLSX_SRCS
    src/XLDocument.cpp
    src/XLWorkbook.cpp
)

add_library(MiniXLSX SHARED ${MLSX_SRCS})

target_include_directories(MiniXLSX PUBLIC include ${PUGIXML_INCLUDE_DIRS})

add_subdirectory(KFZippa)

target_link_libraries(MiniXLSX KFZippa ${PUGIXML_LIBRARIES})

if(MiniXLSX_BUILD_TEST)
    add_subdirectory(tests)
endif(MiniXLSX_BUILD_TEST)

if(MiniXLSX_BUILD_DEMO)
    add_subdirectory(demos)
endif(MiniXLSX_BUILD_DEMO)