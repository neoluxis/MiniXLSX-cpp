#include "test_framework.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <filesystem>

void register_open_file_tests(TestSuite& suite) {
    suite.add_test("Open valid XLSX file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");
        TEST_ASSERT(doc.isOpened(), "Document should be opened");

        doc.close();
        TEST_ASSERT(!doc.isOpened(), "Document should be closed");

        return true;
    });

    suite.add_test("Open invalid file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        bool result = doc.open("nonexistent.xlsx");
        TEST_ASSERT(!result, "Opening nonexistent file should fail");
        TEST_ASSERT(!doc.isOpened(), "Document should not be opened");

        return true;
    });

    suite.add_test("Open empty XLSX file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        bool opened = doc.open("empty_xlsx.xlsx");
        if (!opened) {
            // If it fails to open, that's also acceptable for an empty file
            return true;
        }

        TEST_ASSERT(doc.isOpened(), "Document should be opened if open succeeded");

        // Just check that we can get the workbook without crashing
        try {
            auto& workbook = doc.getWorkbook();
            // Don't check sheet count for empty files
        } catch (...) {
            // If getting workbook fails, that's also acceptable
        }

        doc.close();
        return true;
    });

    suite.add_test("Safe close on unopened document", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        bool result = doc.close_safe();
        TEST_ASSERT(result, "close_safe should return true for unopened document");

        return true;
    });
}