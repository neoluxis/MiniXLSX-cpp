add_executable(test_minixlsx
    test_main.cpp
    test_open_file.cpp
    test_read_data.cpp
    test_picture_detection.cpp
    test_create_file.cpp
)

target_link_libraries(test_minixlsx MiniXLSX)

# Copy test files to build directory
configure_file(${CMAKE_SOURCE_DIR}/assets/test.xlsx ${CMAKE_BINARY_DIR}/test.xlsx COPYONLY)
configure_file(${CMAKE_SOURCE_DIR}/assets/empty_xlsx.xlsx ${CMAKE_BINARY_DIR}/empty_xlsx.xlsx COPYONLY)