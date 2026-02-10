#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include <chrono>
#include <filesystem>
#include <iostream>
#include <stdexcept>

namespace cc::neolux::utils::MiniXLSX
{

    XLDocument::XLDocument() : isOpen(false), workbook(nullptr) {}

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
        workbook = new XLWorkbook(*this);
        if (!workbook->load())
        {
            std::cerr << "Failed to load workbook." << std::endl;
            close();
            return false;
        }
        return true;
    }

    void XLDocument::close()
    {
        if (isOpen)
        {
            delete workbook;
            workbook = nullptr;
            std::filesystem::remove_all(tempDir);
            isOpen = false;
        }
    }

    bool XLDocument::close_safe()
    {
        if(isModified) {
            return false;
        }
        close();
        return true;
    }

    bool XLDocument::isOpened() const
    {
        return isOpen;
    }

    const std::filesystem::path &XLDocument::getTempDir() const
    {
        return tempDir;
    }

    bool XLDocument::saveAs(const std::string &xlsxPath)
    {
        // Implementation for saving the document as a new XLSX file
        // This is a placeholder implementation
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        // Here you would implement the logic to zip the contents of tempDir back into an XLSX file
        // For now, we just simulate success
        this->xlsxPath = xlsxPath;
        isModified = false;
        return true;
    }

    bool XLDocument::save()
    {
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        if (!isModified)
        {
            // No changes to save
            return true;
        }

        // Here you would implement the logic to zip the contents of tempDir back into an XLSX file
        // For now, we just simulate success
        isModified = false;
        return true;
    }

    XLWorkbook& XLDocument::getWorkbook()
    {
        if (!workbook)
        {
            throw std::runtime_error("Workbook not loaded");
        }
        return *workbook;
    }

} // namespace cc::neolux::utils::MiniXLSX