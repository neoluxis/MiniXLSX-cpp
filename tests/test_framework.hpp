#pragma once

#include <iostream>
#include <string>
#include <vector>
#include <functional>

class TestSuite {
public:
    struct TestResult {
        std::string name;
        bool passed;
        std::string message;
    };

    void add_test(const std::string& name, std::function<bool()> test_func) {
        tests_.push_back({name, test_func});
    }

    bool run_all() {
        std::cout << "Running MiniXLSX Tests\n";
        std::cout << "=====================\n\n";

        int passed = 0;
        int failed = 0;
        std::vector<TestResult> results;

        for (const auto& test : tests_) {
            std::cout << "Running: " << test.name << "... ";
            try {
                bool result = test.func();
                if (result) {
                    std::cout << "PASSED\n";
                    passed++;
                    results.push_back({test.name, true, ""});
                } else {
                    std::cout << "FAILED\n";
                    failed++;
                    results.push_back({test.name, false, "Test returned false"});
                }
            } catch (const std::exception& e) {
                std::cout << "FAILED (Exception: " << e.what() << ")\n";
                failed++;
                results.push_back({test.name, false, std::string("Exception: ") + e.what()});
            } catch (...) {
                std::cout << "FAILED (Unknown exception)\n";
                failed++;
                results.push_back({test.name, false, "Unknown exception"});
            }
        }

        std::cout << "\n=====================\n";
        std::cout << "Results: " << passed << " passed, " << failed << " failed\n";

        if (failed > 0) {
            std::cout << "\nFailed tests:\n";
            for (const auto& result : results) {
                if (!result.passed) {
                    std::cout << "  - " << result.name << ": " << result.message << "\n";
                }
            }
        }

        return failed == 0;
    }

private:
    struct Test {
        std::string name;
        std::function<bool()> func;
    };

    std::vector<Test> tests_;
};

#define TEST_ASSERT(condition, message) \
    if (!(condition)) { \
        std::cerr << "Assertion failed: " << message << " at " << __FILE__ << ":" << __LINE__ << std::endl; \
        return false; \
    }

#define TEST_ASSERT_EQUAL(a, b, message) \
    if ((a) != (b)) { \
        std::cerr << "Assertion failed: " << message << " (" << a << " != " << b << ") at " << __FILE__ << ":" << __LINE__ << std::endl; \
        return false; \
    }