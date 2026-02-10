cmake_minimum_required(VERSION 3.11)

project(MiniXLSX)

option(MiniXLSX_BUILD_TEST "Build MiniXLSX test applications" ON)
option(MiniXLSX_BUILD_DEMO "Build MiniXLSX demo applications" ON)
option(MiniXLSX_BUILD_TOOL "Build MiniXLSX tool applications" ON)


set(CMAKE_CXX_STANDARD 20)


message("MiniXLSX: A c++ lib to process xlsx files by LKF")

set(MLSX_SRCS
    src/XLDocument.cpp
    src/XLWorkbook.cpp
    src/XLSheet.cpp
    src/XLCellData.cpp
    src/XLCellPicture.cpp
    src/XLDrawing.cpp
    src/XLTemplate.cpp
)

set(TEMPLCRE_SRCS
    src/tools/XLTemplateCreator.cpp
)

add_library(MiniXLSX SHARED ${MLSX_SRCS})

target_include_directories(MiniXLSX PUBLIC include)

add_subdirectory(KFZippa)

target_link_libraries(MiniXLSX KFZippa)


if(MiniXLSX_BUILD_TOOL)
    add_executable(template_creator ${TEMPLCRE_SRCS})
endif(MiniXLSX_BUILD_TOOL)



if(MiniXLSX_BUILD_TEST)
    add_subdirectory(tests)
endif(MiniXLSX_BUILD_TEST)

if(MiniXLSX_BUILD_DEMO)
    add_subdirectory(demos)
endif(MiniXLSX_BUILD_DEMO)