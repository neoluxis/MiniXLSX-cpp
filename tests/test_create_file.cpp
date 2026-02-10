#include "test_framework.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <filesystem>
#include <string>

void register_create_file_tests(TestSuite& suite) {
    suite.add_test("Create new XLSX file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        std::string testFile = "test_created.xlsx";

        // Clean up any existing file
        if (std::filesystem::exists(testFile)) {
            std::filesystem::remove(testFile);
        }

        TEST_ASSERT(doc.create(testFile), "Failed to create new XLSX file");
        TEST_ASSERT(doc.isOpened(), "Document should be opened after creation");

        auto& workbook = doc.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() == 1, "New file should have 1 sheet");

        std::string sheetName = workbook.getSheetName(0);
        TEST_ASSERT(sheetName == "Sheet1", "First sheet should be named 'Sheet1'");

        auto& sheet = workbook.getSheet(0);
        std::string a1Value = sheet.getCellValue("A1");
        TEST_ASSERT(a1Value.empty(), "A1 should be empty in new file");

        TEST_ASSERT(doc.saveAs(testFile), "Failed to save the created file");

        doc.close();
        TEST_ASSERT(!doc.isOpened(), "Document should be closed");

        // Verify file exists
        TEST_ASSERT(std::filesystem::exists(testFile), "Created file should exist on disk");

        // Clean up
        std::filesystem::remove(testFile);

        return true;
    });

    suite.add_test("Open created XLSX file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        std::string testFile = "test_created2.xlsx";

        // Clean up any existing file
        if (std::filesystem::exists(testFile)) {
            std::filesystem::remove(testFile);
        }

        TEST_ASSERT(doc.create(testFile), "Failed to create new XLSX file");
        TEST_ASSERT(doc.saveAs(testFile), "Failed to save the created file");
        doc.close();

        // Now try to open it
        cc::neolux::utils::MiniXLSX::XLDocument doc2;
        TEST_ASSERT(doc2.open(testFile), "Failed to open the created file");

        auto& workbook = doc2.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() == 1, "Opened file should have 1 sheet");

        doc2.close();

        // Clean up
        std::filesystem::remove(testFile);

        return true;
    });

    suite.add_test("Create and modify file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        std::string testFile = "test_modified.xlsx";

        // Clean up any existing file
        if (std::filesystem::exists(testFile)) {
            std::filesystem::remove(testFile);
        }

        TEST_ASSERT(doc.create(testFile), "Failed to create new XLSX file");

        // File should be marked as modified
        TEST_ASSERT(doc.saveAs(testFile), "Failed to save the created file");
        doc.close();

        // Clean up
        std::filesystem::remove(testFile);

        return true;
    });

    suite.add_test("Create, modify and save XLSX file", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        std::string testFile = "test_modify.xlsx";

        // Clean up any existing file
        if (std::filesystem::exists(testFile)) {
            std::filesystem::remove(testFile);
        }

        TEST_ASSERT(doc.create(testFile), "Failed to create new XLSX file");
        TEST_ASSERT(doc.isOpened(), "Document should be opened after creation");

        auto& workbook = doc.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() == 1, "New file should have 1 sheet");

        auto& sheet = workbook.getSheet(0);

        // Set some cell values
        sheet.setCellValue("A1", "Hello World", "str");
        sheet.setCellValue("B1", "42", "n");
        sheet.setCellValue("C1", "3.14", "n");

        // Save the document
        TEST_ASSERT(doc.saveAs(testFile), "Failed to save the modified file");

        doc.close();
        TEST_ASSERT(!doc.isOpened(), "Document should be closed");

        // Verify file exists
        TEST_ASSERT(std::filesystem::exists(testFile), "Modified file should exist on disk");

        // Re-open and verify the data
        TEST_ASSERT(doc.open(testFile), "Failed to re-open the modified file");

        auto& reopenedWorkbook = doc.getWorkbook();
        auto& reopenedSheet = reopenedWorkbook.getSheet(0);

        TEST_ASSERT(reopenedSheet.getCellValue("A1") == "Hello World", "A1 should contain 'Hello World'");
        TEST_ASSERT(reopenedSheet.getCellValue("B1") == "42", "B1 should contain '42'");
        TEST_ASSERT(reopenedSheet.getCellValue("C1") == "3.14", "C1 should contain '3.14'");

        doc.close();

        // Clean up
        std::filesystem::remove(testFile);

        return true;
    });

    suite.add_test("Fail to create file in invalid location", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        // Try to create in a directory that doesn't exist and we can't write to
        bool result = doc.create("/proc/readonly/test.xlsx");
        TEST_ASSERT(result, "Creating document should succeed (creates temp structure)");
        TEST_ASSERT(doc.isOpened(), "Document should be opened");

        // Now try to save it - this should fail
        bool saveResult = doc.save();
        TEST_ASSERT(!saveResult, "Saving to invalid path should fail");

        doc.close();
        return true;
    });
}