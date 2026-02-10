#include "test_framework.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <string>

void register_read_data_tests(TestSuite& suite) {
    suite.add_test("Read cell A1 from test.xlsx", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() > 0, "Should have sheets");

        auto& sheet = workbook.getSheet(0);
        std::string value = sheet.getCellValue("A1");

        TEST_ASSERT(!value.empty(), "A1 should not be empty");
        TEST_ASSERT(value == "Image  Lists  of  Msr", "A1 should contain expected text");

        doc.close();
        return true;
    });

    suite.add_test("Read cell B2 from test.xlsx", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        auto& sheet = workbook.getSheet(0);
        std::string value = sheet.getCellValue("B2");

        TEST_ASSERT(value == "-5", "B2 should contain '-5'");

        doc.close();
        return true;
    });

    suite.add_test("Read non-existent cell", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        auto& sheet = workbook.getSheet(0);
        std::string value = sheet.getCellValue("Z100");

        TEST_ASSERT(value.empty(), "Non-existent cell should return empty string");

        doc.close();
        return true;
    });

    suite.add_test("Get workbook sheet count", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        int sheetCount = workbook.getSheetCount();

        TEST_ASSERT(sheetCount == 6, "Test file should have 6 sheets");

        doc.close();
        return true;
    });

    suite.add_test("Get sheet names", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        std::string sheetName = workbook.getSheetName(0);

        TEST_ASSERT(!sheetName.empty(), "First sheet should have a name");
        TEST_ASSERT(sheetName == "SO13(DNo.1)MS", "First sheet should be named 'SO13(DNo.1)MS'");

        doc.close();
        return true;
    });
}