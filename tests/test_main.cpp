#include "test_framework.hpp"
#include <iostream>

int main(int argc, char** argv) {
    TestSuite suite;

    // Add all test functions here
    extern void register_open_file_tests(TestSuite& suite);
    extern void register_read_data_tests(TestSuite& suite);
    extern void register_picture_detection_tests(TestSuite& suite);
    extern void register_create_file_tests(TestSuite& suite);

    register_open_file_tests(suite);
    register_read_data_tests(suite);
    register_picture_detection_tests(suite);
    register_create_file_tests(suite);

    bool success = suite.run_all();
    return success ? 0 : 1;
}