find_package(GTest REQUIRED)

add_executable(test_minixlsx
    test_minixlsx_gtest.cpp
    test_style_write_gtest.cpp
)

target_link_libraries(test_minixlsx PRIVATE MiniXLSX GTest::gtest GTest::gtest_main)

# Copy test files to build directory
configure_file(${CMAKE_SOURCE_DIR}/assets/test.xlsx ${CMAKE_BINARY_DIR}/test.xlsx COPYONLY)
configure_file(${CMAKE_SOURCE_DIR}/assets/empty_xlsx.xlsx ${CMAKE_BINARY_DIR}/empty_xlsx.xlsx COPYONLY)

enable_testing()
add_test(NAME minixlsx_tests COMMAND test_minixlsx)