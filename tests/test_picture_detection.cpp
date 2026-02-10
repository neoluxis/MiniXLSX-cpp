#include "test_framework.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include <string>

void register_picture_detection_tests(TestSuite& suite) {
    suite.add_test("Detect picture in G7 of first sheet", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        auto& sheet = workbook.getSheet(0);

        const auto* cell = sheet.getCell("G7");
        TEST_ASSERT(cell != nullptr, "G7 cell should exist");

        TEST_ASSERT(cell->getType() == "picture", "G7 should be a picture cell");

        const auto* picCell = dynamic_cast<const cc::neolux::utils::MiniXLSX::XLCellPicture*>(cell);
        TEST_ASSERT(picCell != nullptr, "G7 should be castable to XLCellPicture");

        std::string imageName = picCell->getImageFileName();
        TEST_ASSERT(!imageName.empty(), "Picture should have a filename");
        TEST_ASSERT(imageName == "image1.jpg", "Picture filename should be 'image1.jpg'");

        std::string fullPath = picCell->getFullPath(doc.getTempDir().string());
        TEST_ASSERT(!fullPath.empty(), "Picture should have a full path");

        doc.close();
        return true;
    });

    suite.add_test("Check G7 in second sheet (no picture)", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() >= 2, "Should have at least 2 sheets");

        auto& sheet = workbook.getSheet(1);
        const auto* cell = sheet.getCell("G7");

        TEST_ASSERT(cell == nullptr, "G7 in second sheet should not exist (no picture)");

        doc.close();
        return true;
    });

    suite.add_test("Check picture in third sheet (SO13(DNo.9)MS)", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        TEST_ASSERT(workbook.getSheetCount() >= 4, "Should have at least 4 sheets");

        auto& sheet = workbook.getSheet(3); // SO13(DNo.9)MS
        const auto* cell = sheet.getCell("G7");

        TEST_ASSERT(cell != nullptr, "G7 in fourth sheet should exist");
        TEST_ASSERT(cell->getType() == "picture", "G7 in fourth sheet should be a picture");

        const auto* picCell = dynamic_cast<const cc::neolux::utils::MiniXLSX::XLCellPicture*>(cell);
        TEST_ASSERT(picCell != nullptr, "Should be castable to XLCellPicture");

        std::string imageName = picCell->getImageFileName();
        TEST_ASSERT(imageName == "image179.jpg", "Picture should be 'image179.jpg'");

        doc.close();
        return true;
    });

    suite.add_test("Check cell types for data cells", []() -> bool {
        cc::neolux::utils::MiniXLSX::XLDocument doc;

        TEST_ASSERT(doc.open("test.xlsx"), "Failed to open test.xlsx");

        auto& workbook = doc.getWorkbook();
        auto& sheet = workbook.getSheet(0);

        const auto* cellA1 = sheet.getCell("A1");
        TEST_ASSERT(cellA1 != nullptr, "A1 should exist");
        std::string typeA1 = cellA1->getType();
        TEST_ASSERT(!typeA1.empty(), "A1 should have a type");

        const auto* cellB2 = sheet.getCell("B2");
        TEST_ASSERT(cellB2 != nullptr, "B2 should exist");
        std::string typeB2 = cellB2->getType();
        TEST_ASSERT(!typeB2.empty(), "B2 should have a type");

        doc.close();
        return true;
    });
}