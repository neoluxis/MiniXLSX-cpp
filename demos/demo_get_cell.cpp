#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <iostream>
#include <string>

int main(int argc, const char** argv)
{
    if (argc < 2)
    {
        std::cout << "Usage: demo_get_cell <path_to_xlsx>" << std::endl;
        return 1;
    }
    std::string xlsxPath = argv[1];

    cc::neolux::utils::MiniXLSX::XLDocument doc;

    // Test opening
    if (doc.open(xlsxPath))
    {
        std::cout << "Opened successfully. Temp dir: " << doc.getTempDir() << std::endl;

        // Test workbook
        auto& workbook = doc.getWorkbook();
        std::cout << "Number of sheets: " << workbook.getSheetCount() << std::endl;

        if (workbook.getSheetCount() > 0)
        {
            auto& sheet = workbook.getSheet(0);
            std::cout << "Sheet name: " << sheet.getName() << std::endl;

            // Test reading some cells
            std::string a1 = sheet.getCellValue("A1");
            std::string b2 = sheet.getCellValue("B2");
            std::cout << "Original A1: '" << a1 << "'" << std::endl;
            std::cout << "Original B2: '" << b2 << "'" << std::endl;

            // Test writing to cells
            sheet.setCellValue("A1", "Modified A1", "str");
            sheet.setCellValue("B2", "999", "n");
            sheet.setCellValue("C3", "New Cell", "str");

            std::cout << "After modification:" << std::endl;
            std::cout << "A1: '" << sheet.getCellValue("A1") << "'" << std::endl;
            std::cout << "B2: '" << sheet.getCellValue("B2") << "'" << std::endl;
            std::cout << "C3: '" << sheet.getCellValue("C3") << "'" << std::endl;

            // Save the changes
            std::string outputPath = "modified_" + xlsxPath;
            if (doc.saveAs(outputPath))
            {
                std::cout << "Saved modified file to: " << outputPath << std::endl;
            }
            else
            {
                std::cout << "Failed to save modified file." << std::endl;
            }
        }

        // Close
        doc.close();
        std::cout << "Closed." << std::endl;
    }
    else
    {
        std::cout << "Failed to open." << std::endl;
    }

    return 0;
}