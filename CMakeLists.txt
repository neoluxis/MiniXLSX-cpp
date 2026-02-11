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
add_subdirectory(3rdparty/OpenXLSX)

target_link_libraries(MiniXLSX PUBLIC KFZippa)
target_link_libraries(MiniXLSX PUBLIC OpenXLSX::OpenXLSX)

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
    add_subdirectory(tests)
endif(MiniXLSX_BUILD_TEST)

if(MiniXLSX_BUILD_DEMO)
    add_subdirectory(demos)
endif(MiniXLSX_BUILD_DEMO)