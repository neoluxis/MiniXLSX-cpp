#include "cc/neolux/utils/MiniXLSX/XLWorkbook.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <filesystem>
#include <iostream>
#include <pugixml.hpp>

namespace cc::neolux::utils::MiniXLSX
{

    XLWorkbook::XLWorkbook(XLDocument& doc) : document(&doc)
    {
    }

    XLWorkbook::~XLWorkbook()
    {
    }

    bool XLWorkbook::load()
    {
        if (!document->isOpened())
        {
            std::cerr << "Document is not open." << std::endl;
            return false;
        }

        std::filesystem::path workbookPath = document->getTempDir() / "xl" / "workbook.xml";
        
        pugi::xml_document doc;
        pugi::xml_parse_result result = doc.load_file(workbookPath.c_str());
        if (!result)
        {
            std::cerr << "Failed to parse workbook.xml: " << result.description() << std::endl;
            return false;
        }

        auto workbook = doc.child("workbook");
        if (!workbook)
        {
            std::cerr << "No workbook element found." << std::endl;
            return false;
        }

        auto sheets = workbook.child("sheets");
        if (!sheets)
        {
            std::cerr << "No sheets element found." << std::endl;
            return false;
        }

        for (auto sheet : sheets.children("sheet"))
        {
            std::string name = sheet.attribute("name").value();
            std::string sheetId = sheet.attribute("sheetId").value();
            std::string rId = sheet.attribute("r:id").value();

            sheetNames.push_back(name);
            sheetIds.push_back(sheetId);
            rIds.push_back(rId);
        }

        return true;
    }

    size_t XLWorkbook::getSheetCount() const
    {
        return sheetNames.size();
    }

    const std::string& XLWorkbook::getSheetName(size_t index) const
    {
        if (index >= sheetNames.size())
        {
            static std::string empty;
            return empty;
        }
        return sheetNames[index];
    }

} // namespace cc::neolux::utils::MiniXLSX