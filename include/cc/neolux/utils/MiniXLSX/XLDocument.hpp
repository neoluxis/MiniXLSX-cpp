#pragma once

#include <string>
#include <filesystem>
#include <map>
#include <vector>
#include "XLWorkbook.hpp"
#include "OpenXLSXWrapper.hpp"

namespace cc::neolux::utils::MiniXLSX
{

class XLDocument
{
private:
    std::filesystem::path tempDir;
    bool isOpen;
    bool isModified;
    std::string xlsxPath;
    XLWorkbook* workbook;
    std::unique_ptr<OpenXLSXWrapper> oxwrapper;

public:
    XLDocument();
    ~XLDocument();

    // Access to OpenXLSX wrapper if available
    OpenXLSXWrapper* getWrapper() const;

    /**
     * @brief Opens an XLSX file by unzipping it to a temporary directory.
     * @param xlsxPath The path to the XLSX file.
     * @return true if successful, false otherwise.
     */
    bool open(const std::string& xlsxPath);

    /**
     * @brief Creates a new XLSX file with basic structure.
     * @param xlsxPath The path where the new XLSX file will be created.
     * @return true if successful, false otherwise.
     */
    bool create(const std::string& xlsxPath);

    /**
     * @brief Closes the document by cleaning up the temporary directory.
     */
    void close();

    /**
     * @brief Close the document safely.
     * @return true if the document was closed, false if it was modified and could not be closed.
     */
    bool close_safe();

    /**
     * @brief Checks if the document is currently open.
     * @return true if open, false otherwise.
     */
    bool isOpened() const;

    /**
     * @brief Gets the temporary directory path where the XLSX is extracted.
     * @return The temporary directory path.
     */
    const std::filesystem::path& getTempDir() const;

    /**
     * @brief Gets the workbook associated with this document.
     * @return Reference to the workbook.
     */
    XLWorkbook& getWorkbook();


    /**
     * @brief Marks the document as modified.
     * @param xlsxPath The path to the XLSX file.
     * @return true if successful, false otherwise.
     */
    bool saveAs(const std::string& xlsxPath);

    /**
     * @brief Saves the document if it has been modified.
     * @return true if successful, false otherwise.
     */
    bool save();

    /**
     * @brief Marks the document as modified.
     */
    void markModified();
};

} // namespace cc::neolux::utils::MiniXLSX