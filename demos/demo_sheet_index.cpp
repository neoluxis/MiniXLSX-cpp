#include <iostream>
#include <string>
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"

int main(int argc, char** argv)
{
    if (argc < 2) {
        std::cerr << "Usage: " << argv[0] << " <xlsx_file_path>" << std::endl;
        std::cerr << "Example: " << argv[0] << " test.xlsx" << std::endl;
        return 1;
    }

    std::string path = argv[1];
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper wrapper;

    if (!wrapper.open(path)) {
        std::cerr << "Failed to open input xlsx: " << path << std::endl;
        return 2;
    }

    std::cout << "Successfully opened: " << path << std::endl;
    std::cout << "Number of sheets: " << wrapper.sheetCount() << std::endl;

    // Demonstrate sheet name to index lookup
    std::cout << "\n=== Sheet Name to Index Lookup Demo ===" << std::endl;

    // List all sheets with their indices
    for (unsigned int i = 0; i < wrapper.sheetCount(); ++i) {
        std::string name = wrapper.sheetName(i);
        std::cout << "Sheet " << i << ": '" << name << "'" << std::endl;
    }

    // Demonstrate sheetIndex function
    std::cout << "\n=== Testing sheetIndex function ===" << std::endl;

    // Test with first sheet
    if (wrapper.sheetCount() > 0) {
        std::string firstSheetName = wrapper.sheetName(0);
        auto indexOpt = wrapper.sheetIndex(firstSheetName);
        if (indexOpt.has_value()) {
            std::cout << "Sheet '" << firstSheetName << "' has index: " << indexOpt.value() << std::endl;
        } else {
            std::cout << "ERROR: Could not find index for sheet '" << firstSheetName << "'" << std::endl;
        }
    }

    // Test with invalid sheet name
    auto invalidIndexOpt = wrapper.sheetIndex("NonExistentSheet");
    if (!invalidIndexOpt.has_value()) {
        std::cout << "Correctly returned nullopt for non-existent sheet 'NonExistentSheet'" << std::endl;
    } else {
        std::cout << "ERROR: Should have returned nullopt for non-existent sheet" << std::endl;
    }

    // Demonstrate round-trip functionality
    std::cout << "\n=== Round-trip Test (Name -> Index -> Name) ===" << std::endl;
    bool allPassed = true;
    for (unsigned int i = 0; i < wrapper.sheetCount(); ++i) {
        std::string originalName = wrapper.sheetName(i);
        auto indexOpt = wrapper.sheetIndex(originalName);
        if (indexOpt.has_value()) {
            std::string retrievedName = wrapper.sheetName(indexOpt.value());
            if (originalName == retrievedName) {
                std::cout << "✓ Sheet " << i << ": '" << originalName << "' -> " << indexOpt.value() << " -> '" << retrievedName << "'" << std::endl;
            } else {
                std::cout << "✗ Sheet " << i << ": Mismatch - Original: '" << originalName << "', Retrieved: '" << retrievedName << "'" << std::endl;
                allPassed = false;
            }
        } else {
            std::cout << "✗ Sheet " << i << ": Could not find index for '" << originalName << "'" << std::endl;
            allPassed = false;
        }
    }

    if (allPassed) {
        std::cout << "\n✓ All round-trip tests passed!" << std::endl;
    } else {
        std::cout << "\n✗ Some round-trip tests failed!" << std::endl;
    }

    wrapper.close();
    std::cout << "\nDemo completed successfully!" << std::endl;
    return 0;
}