#pragma once

#include <string>
#include <vector>

namespace cc::neolux::utils::MiniXLSX
{
    class XLDocument;

    class XLWorkbook
    {
    private:
        XLDocument* document;
        std::vector<std::string> sheetNames;
        std::vector<std::string> sheetIds;
        std::vector<std::string> rIds;
    public:
        XLWorkbook(XLDocument& doc);
        ~XLWorkbook();

        /**
         * @brief Loads the workbook data from the document.
         * @return true if successful, false otherwise.
         */
        bool load();

        /**
         * @brief Gets the number of sheets in the workbook.
         * @return The number of sheets.
         */
        size_t getSheetCount() const;

        /**
         * @brief Gets the name of the sheet at the given index.
         * @param index The index of the sheet.
         * @return The sheet name.
         */
        const std::string& getSheetName(size_t index) const;
    };
    
} // namespace cc::neolux::utils::MiniXLSX