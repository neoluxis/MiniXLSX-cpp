#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include <iostream>
#include <string>

int main(int argc, const char** argv)
{
    if (argc < 2)
    {
        std::cout << "Usage: demo_open_close <path_to_xlsx>" << std::endl;
        return 1;
    }
    std::string xlsxPath = argv[1];

    cc::neolux::utils::MiniXLSX::XLDocument doc;
    
    // Test opening
    if (doc.open(xlsxPath))
    {
        std::cout << "Opened successfully. Temp dir: " << doc.getTempDir() << std::endl;
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

            // Check G7 for picture
            const auto* cellG7 = sheet.getCell("G7");
            if (cellG7)
            {
                std::cout << "  Cell G7: type='" << cellG7->getType() << "', value='" << cellG7->getValue() << "'" << std::endl;
                if (cellG7->getType() == "picture")
                {
                    // Cast to XLCellPicture
                    const auto* picCell = dynamic_cast<const cc::neolux::utils::MiniXLSX::XLCellPicture*>(cellG7);
                    if (picCell)
                    {
                        std::cout << "    Picture name: '" << picCell->getImageFileName() << "'" << std::endl;
                        std::cout << "    Full path: '" << picCell->getFullPath(doc.getTempDir().string()) << "'" << std::endl;
                    }
                }
            }
            else
            {
                std::cout << "  Cell G7: not found" << std::endl;
            }
        }

        // Close
        doc.close();
        std::cout << "Closed. Is open: " << (doc.isOpened() ? "Yes" : "No") << std::endl;
    }
    else
    {
        std::cout << "Failed to open." << std::endl;
    }

    return 0;
}