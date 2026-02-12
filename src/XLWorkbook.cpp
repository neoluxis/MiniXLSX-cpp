#include "cc/neolux/utils/MiniXLSX/XLWorkbook.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>

namespace cc::neolux::utils::MiniXLSX
{

    XLWorkbook::XLWorkbook(XLDocument& doc) : document(&doc)
    {
    }

    XLWorkbook::~XLWorkbook()
    {
        for (auto sheet : sheets)
        {
            delete sheet;
        }
        sheets.clear();
    }

    bool XLWorkbook::load()
    {
        if (!document->isOpened())
        {
            std::cerr << "Document is not open." << std::endl;
            return false;
        }

        // 若 OpenXLSX 封装可用，则基于其构建工作表列表
        try {
            OpenXLSXWrapper* w = document->getWrapper();
            XLPictureReader* pr = document->getPictureReader();
            if (w && w->isOpen()) {
                unsigned int cnt = w->sheetCount();
                for (unsigned int i = 0; i < cnt; ++i) {
                    std::string name = w->sheetName(i);
                    XLSheet* xlSheet = new XLSheet(*this, w, pr, i);
                    // 不调用 load()，由封装提供单元格访问
                    sheets.push_back(xlSheet);
                    sheetNames.push_back(name);
                }
                return true;
            }
        } catch(...) {}

        std::filesystem::path workbookPath = document->getTempDir() / "xl" / "workbook.xml";
        std::ifstream file(workbookPath);
        if (!file.is_open())
        {
            std::cerr << "Failed to open workbook.xml" << std::endl;
            return false;
        }

        std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
        file.close();

        // Simple parsing for <sheet name="..." sheetId="..." r:id="..."/>
        size_t pos = 0;
        while ((pos = content.find("<sheet ", pos)) != std::string::npos)
        {
            size_t endPos = content.find("/>", pos);
            if (endPos == std::string::npos) break;

            std::string sheetTag = content.substr(pos, endPos - pos + 2);

            // Extract name
            size_t nameStart = sheetTag.find("name=\"");
            if (nameStart != std::string::npos)
            {
                nameStart += 6;
                size_t nameEnd = sheetTag.find("\"", nameStart);
                if (nameEnd != std::string::npos)
                {
                    std::string name = sheetTag.substr(nameStart, nameEnd - nameStart);
                    sheetNames.push_back(name);
                }
            }

            // Extract sheetId
            size_t idStart = sheetTag.find("sheetId=\"");
            if (idStart != std::string::npos)
            {
                idStart += 9;
                size_t idEnd = sheetTag.find("\"", idStart);
                if (idEnd != std::string::npos)
                {
                    std::string id = sheetTag.substr(idStart, idEnd - idStart);
                    sheetIds.push_back(id);
                }
            }

            // Extract r:id
            size_t rStart = sheetTag.find("r:id=\"");
            if (rStart != std::string::npos)
            {
                rStart += 6;
                size_t rEnd = sheetTag.find("\"", rStart);
                if (rEnd != std::string::npos)
                {
                    std::string rId = sheetTag.substr(rStart, rEnd - rStart);
                    rIds.push_back(rId);
                }
            }

            pos = endPos + 2;
        }

        // Create sheets
        for (size_t i = 0; i < sheetNames.size(); ++i)
        {
            XLSheet* xlSheet = new XLSheet(*this, sheetNames[i], sheetIds[i], rIds[i]);
            if (!xlSheet->load())
            {
                std::cerr << "Failed to load sheet: " << sheetNames[i] << std::endl;
                delete xlSheet;
                // Continue or return false?
            }
            else
            {
                sheets.push_back(xlSheet);
            }
        }

        return true;
    }

    bool XLWorkbook::save()
    {
        for (XLSheet* sheet : sheets)
        {
            if (!sheet->save())
            {
                std::cerr << "Failed to save sheet: " << sheet->getName() << std::endl;
                return false;
            }
        }
        return true;
    }

    size_t XLWorkbook::getSheetCount() const
    {
        // 优先使用 OpenXLSX 封装
        try {
            OpenXLSXWrapper* w = document->getWrapper();
            if (w && w->isOpen()) return static_cast<size_t>(w->sheetCount());
        } catch(...) {}
        return sheets.size();
    }

    const std::string& XLWorkbook::getSheetName(size_t index) const
    {
        try {
            OpenXLSXWrapper* w = document->getWrapper();
            if (w && w->isOpen()) {
                std::string name = w->sheetName(static_cast<unsigned int>(index));
                // return as static copy to satisfy reference return — small hack for compatibility
                static thread_local std::string tmp;
                tmp = name;
                return tmp;
            }
        } catch(...) {}
        if (index >= sheetNames.size())
        {
            static std::string empty;
            return empty;
        }
        return sheetNames[index];
    }

    XLSheet& XLWorkbook::getSheet(size_t index)
    {
        if (index >= sheets.size())
        {
            throw std::out_of_range("Sheet index out of range");
        }
        return *sheets[index];
    }

    XLDocument& XLWorkbook::getDocument() const
    {
        return *document;
    }

} // namespace cc::neolux::utils::MiniXLSX