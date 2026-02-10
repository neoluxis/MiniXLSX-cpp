#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <iostream>
#include <string>

int main(int argc, const char** argv)
{
    if (argc < 2)
    {
        std::cout << "Usage: demo_create <path_to_new_xlsx>" << std::endl;
        return 1;
    }
    std::string xlsxPath = argv[1];

    cc::neolux::utils::MiniXLSX::XLDocument doc;

    // Test creating a new XLSX file
    if (doc.create(xlsxPath))
    {
        std::cout << "Created successfully. Temp dir: " << doc.getTempDir() << std::endl;
        std::cout << "Is open: " << (doc.isOpened() ? "Yes" : "No") << std::endl;

        // Test workbook
        auto& workbook = doc.getWorkbook();
        std::cout << "Number of sheets: " << workbook.getSheetCount() << std::endl;
        for (size_t i = 0; i < workbook.getSheetCount(); ++i)
        {
            std::cout << "Sheet " << i << ": " << workbook.getSheetName(i) << std::endl;
            auto& sheet = workbook.getSheet(i);
            // Test some cells
            std::string a1 = sheet.getCellValue("A1");
            std::string b2 = sheet.getCellValue("B2");
            std::cout << "  Cell A1: '" << a1 << "' (length: " << a1.length() << ")" << std::endl;
            std::cout << "  Cell B2: '" << b2 << "' (length: " << b2.length() << ")" << std::endl;
        }

        // Save the document
        if (doc.saveAs(xlsxPath))
        {
            std::cout << "Saved successfully to: " << xlsxPath << std::endl;
        }
        else
        {
            std::cout << "Failed to save." << std::endl;
        }

        // Close
        doc.close();
        std::cout << "Closed. Is open: " << (doc.isOpened() ? "Yes" : "No") << std::endl;
    }
    else
    {
        std::cout << "Failed to create." << std::endl;
    }

    return 0;
}