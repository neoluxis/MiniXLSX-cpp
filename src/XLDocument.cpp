#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include "cc/neolux/utils/MiniXLSX/XLTemplate.hpp"
#include <chrono>
#include <filesystem>
#include <iostream>
#include <fstream>
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

    bool XLDocument::create(const std::string& xlsxPath)
    {
        if (isOpen)
        {
            close();
        }

        XLTemplate::createTemplate(xlsxPath);
        return open(xlsxPath);
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
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        // Save all changes to XML files
        if (!workbook->save())
        {
            std::cerr << "Failed to save workbook changes." << std::endl;
            return false;
        }

        // Use KFZippa to zip the tempDir contents to the xlsxPath
        if (!KFZippa::zip(tempDir.string(), xlsxPath))
        {
            std::cerr << "Failed to zip contents to XLSX file: " << xlsxPath << std::endl;
            return false;
        }

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

        // Save all changes to XML files
        if (!workbook->save())
        {
            std::cerr << "Failed to save workbook changes." << std::endl;
            return false;
        }

        // Use KFZippa to zip the tempDir contents to the xlsxPath
        if (!KFZippa::zip(tempDir.string(), xlsxPath))
        {
            std::cerr << "Failed to zip contents to XLSX file: " << xlsxPath << std::endl;
            return false;
        }

        isModified = false;
        return true;
    }

    void XLDocument::markModified()
    {
        isModified = true;
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