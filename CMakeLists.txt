cmake_minimum_required(VERSION 3.11)

project(MiniXLSX)

option(MiniXLSX_BUILD_TEST "Build MiniXLSX test applications" OFF)
option(MiniXLSX_BUILD_DEMO "Build MiniXLSX demo applications" OFF)
option(MiniXLSX_BUILD_TOOL "Build MiniXLSX tool applications" OFF)


set(CMAKE_CXX_STANDARD 20)


message("MiniXLSX: A c++ lib to process xlsx files by LKF")

set(MLSX_SRCS
    src/XLDocument.cpp
    src/XLWorkbook.cpp
    src/XLSheet.cpp
    src/XLCellData.cpp
    src/XLCellPicture.cpp
    src/XLPictureReader.cpp
    src/XLTemplate.cpp
    src/OpenXLSXWrapper.cpp
    src/MiniXLSX.cpp
)

set(TEMPLCRE_SRCS
    src/tools/XLTemplateCreator.cpp
)

add_library(MiniXLSX SHARED ${MLSX_SRCS})

target_include_directories(MiniXLSX PUBLIC include)

add_subdirectory(3rdparty/KFZippa)

# 在 Windows/Mingw 下使用 OpenXLSX 共享库，避免静态库符号重复
if (CMAKE_SYSTEM_NAME STREQUAL "Windows" OR MINGW)
    set(OPENXLSX_LIBRARY_TYPE "SHARED" CACHE STRING "" FORCE)
endif()
add_subdirectory(3rdparty/OpenXLSX)

target_link_libraries(MiniXLSX PUBLIC KFZippa)
target_link_libraries(MiniXLSX PRIVATE OpenXLSX::OpenXLSX)

# Find system-installed pugixml via pkg-config and link it
find_package(PkgConfig REQUIRED)
pkg_check_modules(PUGIXML QUIET pugixml)
if (NOT PUGIXML_FOUND)
    pkg_check_modules(PUGIXML QUIET pugi)
endif()
if (PUGIXML_FOUND)
    target_include_directories(MiniXLSX PUBLIC ${PUGIXML_INCLUDE_DIRS})
    target_link_directories(MiniXLSX PUBLIC ${PUGIXML_LIBRARY_DIRS})
    target_link_libraries(MiniXLSX PUBLIC ${PUGIXML_LIBRARIES})
    target_compile_options(MiniXLSX PUBLIC ${PUGIXML_CFLAGS_OTHER})
else()
    message(FATAL_ERROR "pugixml not found. Please install system pugixml (pkg-config entry 'pugixml' or 'pugi').")
endif()


if(MiniXLSX_BUILD_TOOL)
    add_executable(template_creator ${TEMPLCRE_SRCS})
endif(MiniXLSX_BUILD_TOOL)



if(MiniXLSX_BUILD_TEST)
    enable_testing()
    add_subdirectory(tests)
endif(MiniXLSX_BUILD_TEST)

if(MiniXLSX_BUILD_DEMO)
    add_subdirectory(demos)
endif(MiniXLSX_BUILD_DEMO)