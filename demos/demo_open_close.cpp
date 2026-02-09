#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <iostream>
#include "string"

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