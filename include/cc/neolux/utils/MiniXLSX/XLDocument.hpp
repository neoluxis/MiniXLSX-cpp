#pragma once

#include <string>
#include <filesystem>

namespace cc::neolux::utils::MiniXLSX
{

class XLDocument
{
private:
    std::filesystem::path tempDir;
    bool isOpen;

public:
    XLDocument();
    ~XLDocument();

    /**
     * @brief Opens an XLSX file by unzipping it to a temporary directory.
     * @param xlsxPath The path to the XLSX file.
     * @return true if successful, false otherwise.
     */
    bool open(const std::string& xlsxPath);

    /**
     * @brief Closes the document by cleaning up the temporary directory.
     */
    void close();

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
};

} // namespace cc::neolux::utils::MiniXLSX