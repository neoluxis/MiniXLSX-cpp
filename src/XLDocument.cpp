#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include <chrono>
#include <filesystem>
#include <iostream>

namespace cc::neolux::utils::MiniXLSX
{

    XLDocument::XLDocument() : isOpen(false) {}

    XLDocument::~XLDocument()
    {
        close();
    }

    bool XLDocument::open(const std::string &xlsxPath)
    {
        if (isOpen)
        {
            close();
        }

        // Create a unique temporary directory
        auto now = std::chrono::system_clock::now();
        auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
        std::string tempDirName = "MiniXLSX_" + std::to_string(timestamp);
        tempDir = std::filesystem::temp_directory_path() / tempDirName;

        // Create the directory
        if (!std::filesystem::create_directory(tempDir))
        {
            std::cerr << "Failed to create temporary directory: " << tempDir << std::endl;
            return false;
        }

        // Unzip the XLSX file
        if (!KFZippa::unzip(xlsxPath, tempDir.string()))
        {
            std::cerr << "Failed to unzip XLSX file: " << xlsxPath << std::endl;
            std::filesystem::remove_all(tempDir); // Clean up on failure
            return false;
        }

        isOpen = true;
        return true;
    }

    void XLDocument::close()
    {
        if (isOpen)
        {
            std::filesystem::remove_all(tempDir);
            isOpen = false;
        }
    }

    bool XLDocument::isOpened() const
    {
        return isOpen;
    }

    const std::filesystem::path &XLDocument::getTempDir() const
    {
        return tempDir;
    }

} // namespace cc::neolux::utils::MiniXLSX