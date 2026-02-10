#pragma once

#include <string>
#include <map>
#include <vector>
#include <memory>
#include "XLCell.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    class XLWorkbook;

    class XLSheet
    {
    private:
        XLWorkbook* workbook;
        std::string name;
        std::string sheetId;
        std::string rId;
        std::map<std::string, std::unique_ptr<XLCell>> cells;
        std::vector<std::string> sharedStrings;

    public:
        XLSheet(XLWorkbook& wb, const std::string& n, const std::string& sid, const std::string& rid);
        ~XLSheet();

        /**
         * @brief Loads the sheet data from the document.
         * @return true if successful, false otherwise.
         */
        bool load();

        /**
         * @brief Gets the sheet name.
         * @return The sheet name.
         */
        const std::string& getName() const;

        /**
         * @brief Gets the cell by reference (e.g., "A1").
         * @param ref The cell reference.
         * @return Pointer to the cell, or nullptr if not found.
         */
        const XLCell* getCell(const std::string& ref) const;

        /**
         * @brief Gets the value of a cell by reference (e.g., "A1").
         * @param ref The cell reference.
         * @return The cell value, or empty string if not found.
         */
        std::string getCellValue(const std::string& ref) const;

        /**
         * @brief Sets the value of a cell by reference (e.g., "A1").
         * @param ref The cell reference.
         * @param value The value to set.
         * @param type The cell type ("str" for string, "n" for number, etc.).
         */
        void setCellValue(const std::string& ref, const std::string& value, const std::string& type = "str");

        /**
         * @brief Saves the sheet data back to the XML file.
         * @return true if successful, false otherwise.
         */
        bool save();

        // Iterators
        using iterator = std::map<std::string, std::unique_ptr<XLCell>>::const_iterator;
        iterator begin() const { return cells.begin(); }
        iterator end() const { return cells.end(); }

        // Utility function
        static std::string columnNumberToLetter(int col);
    };

} // namespace cc::neolux::utils::MiniXLSX