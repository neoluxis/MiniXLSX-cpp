#include "cc/neolux/utils/MiniXLSX/XLSheet.hpp"
#include "cc/neolux/utils/MiniXLSX/XLWorkbook.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellData.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDrawing.hpp"
#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>
#include <memory>
#include <map>
#include <algorithm>
#include <pugixml.hpp>

namespace cc::neolux::utils::MiniXLSX
{

    std::string XLSheet::columnNumberToLetter(int col)
    {
        std::string result;
        while (col >= 0)
        {
            result = char('A' + (col % 26)) + result;
            col = col / 26 - 1;
        }
        return result;
    }

    XLSheet::XLSheet(XLWorkbook& wb, const std::string& n, const std::string& sid, const std::string& rid)
        : workbook(&wb), name(n), sheetId(sid), rId(rid)
    {
    }

    XLSheet::~XLSheet()
    {
    }

    bool XLSheet::load()
    {
        namespace fs = std::filesystem;
        fs::path temp = workbook->getDocument().getTempDir();

        // Load shared strings using pugixml
        fs::path sharedStringsPath = temp / "xl" / "sharedStrings.xml";
        if (fs::exists(sharedStringsPath))
        {
            pugi::xml_document sdoc;
            pugi::xml_parse_result sres = sdoc.load_file(sharedStringsPath.c_str());
            if (sres)
            {
                pugi::xml_node sst = sdoc.child("sst");
                for (pugi::xml_node si : sst.children("si"))
                {
                    // Prefer direct <t> child but handle rich text
                    pugi::xml_node tnode = si.child("t");
                    if (tnode)
                    {
                        sharedStrings.push_back(tnode.text().get());
                    }
                    else
                    {
                        std::string acc;
                        for (pugi::xpath_node xpath_node : si.select_nodes(".//t"))
                        {
                            acc += xpath_node.node().text().get();
                        }
                        if (!acc.empty()) sharedStrings.push_back(acc);
                    }
                }
            }
        }

        // Find the sheet file path using rId via workbook rels
        fs::path relsPath = temp / "xl" / "_rels" / "workbook.xml.rels";
        pugi::xml_document relsDoc;
        if (!fs::exists(relsPath) || !relsDoc.load_file(relsPath.c_str()))
        {
            std::cerr << "Failed to open workbook.xml.rels" << std::endl;
            return false;
        }

        std::string target;
        pugi::xml_node relsRoot = relsDoc.child("Relationships");
        for (pugi::xml_node rel : relsRoot.children("Relationship"))
        {
            pugi::xml_attribute idAttr = rel.attribute("Id");
            if (idAttr && idAttr.value() == rId)
            {
                pugi::xml_attribute targetAttr = rel.attribute("Target");
                if (targetAttr) { target = targetAttr.value(); break; }
            }
        }

        if (target.empty())
        {
            std::cerr << "Target not found for rId: " << rId << std::endl;
            return false;
        }

        fs::path sheetPath = temp / "xl" / target;
        pugi::xml_document sheetDoc;
        if (!fs::exists(sheetPath) || !sheetDoc.load_file(sheetPath.c_str()))
        {
            std::cerr << "Failed to open sheet: " << sheetPath << std::endl;
            return false;
        }

        // Parse sheetData
        pugi::xml_node worksheet = sheetDoc.child("worksheet");
        pugi::xml_node sheetData = worksheet.child("sheetData");
        if (!sheetData)
        {
            std::cerr << "No sheetData found." << std::endl;
            return false;
        }

        for (pugi::xml_node row : sheetData.children("row"))
        {
            for (pugi::xml_node c : row.children("c"))
            {
                std::string ref = c.attribute("r").as_string();
                std::string type = c.attribute("t").as_string();
                std::string value;

                pugi::xml_node v = c.child("v");
                if (v)
                {
                    value = v.text().get();
                }
                else
                {
                    pugi::xml_node is = c.child("is");
                    if (is)
                    {
                        pugi::xml_node t = is.child("t");
                        if (t) value = t.text().get();
                    }
                }

                if (!ref.empty())
                {
                    cells[ref] = std::make_unique<XLCellData>(ref, value, type.empty() ? "n" : type, sharedStrings);
                }
            }
        }

        // Parse drawings (if any)
        pugi::xml_node drawingNode = worksheet.child("drawing");
        if (drawingNode)
        {
            pugi::xml_attribute ridAttr = drawingNode.attribute("r:id");
            if (ridAttr)
            {
                std::string drawingRId = ridAttr.value();

                fs::path sheetRelsPath = temp / "xl" / "worksheets" / "_rels" / (fs::path(target).filename().string() + ".rels");
                pugi::xml_document sheetRelsDoc;
                if (fs::exists(sheetRelsPath) && sheetRelsDoc.load_file(sheetRelsPath.c_str()))
                {
                    std::string drawingTarget;
                    pugi::xml_node sroot = sheetRelsDoc.child("Relationships");
                    for (pugi::xml_node rel : sroot.children("Relationship"))
                    {
                        if (std::string(rel.attribute("Id").as_string()) == drawingRId)
                        {
                            drawingTarget = rel.attribute("Target").as_string();
                            break;
                        }
                    }

                    if (!drawingTarget.empty())
                    {
                        fs::path drawingPath = (temp / "xl" / "worksheets" / drawingTarget).lexically_normal();
                        fs::path drawingRelsPath = drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");

                        XLDrawing drawing(drawingPath.string(), drawingRelsPath.string());
                        if (drawing.load())
                        {
                            for (const auto& pic : drawing.getPictures())
                            {
                                cells[pic.ref] = std::make_unique<XLCellPicture>(pic.ref, pic.imageFileName, pic.relativePath);
                            }
                        }
                    }
                }
            }
        }

        return true;
    }

    const std::string& XLSheet::getName() const
    {
        return name;
    }

    const XLCell* XLSheet::getCell(const std::string& ref) const
    {
        auto it = cells.find(ref);
        if (it != cells.end())
        {
            return it->second.get();
        }
        return nullptr;
    }

    std::string XLSheet::getCellValue(const std::string& ref) const
    {
        const XLCell* cell = getCell(ref);
        if (cell)
        {
            return cell->getValue();
        }
        return "";
    }

    void XLSheet::setCellValue(const std::string& ref, const std::string& value, const std::string& type)
    {
        auto it = cells.find(ref);
        if (it != cells.end())
        {
            // Cell exists, update it if it's XLCellData
            XLCellData* dataCell = dynamic_cast<XLCellData*>(it->second.get());
            if (dataCell)
            {
                dataCell->setValue(value);
                dataCell->setType(type);
            }
        }
        else
        {
            // Cell doesn't exist, create it
            cells[ref] = std::make_unique<XLCellData>(ref, value, type, sharedStrings);
        }
        
        // Mark document as modified
        workbook->getDocument().markModified();
    }

    bool XLSheet::save()
    {
        // Find the sheet file path using rId
        std::filesystem::path relsPath = workbook->getDocument().getTempDir() / "xl" / "_rels" / "workbook.xml.rels";
        
        std::ifstream relsFile(relsPath);
        if (!relsFile.is_open())
        {
            std::cerr << "Failed to open workbook.xml.rels" << std::endl;
            return false;
        }

        std::string relsContent((std::istreambuf_iterator<char>(relsFile)), std::istreambuf_iterator<char>());
        relsFile.close();

        std::string target;
        size_t relPos = 0;
        while ((relPos = relsContent.find("<Relationship ", relPos)) != std::string::npos)
        {
            size_t relEnd = relsContent.find("/>", relPos);
            if (relEnd == std::string::npos) break;

            std::string relTag = relsContent.substr(relPos, relEnd - relPos + 2);

            size_t idStart = relTag.find("Id=\"");
            if (idStart != std::string::npos)
            {
                idStart += 4;
                size_t idEnd = relTag.find("\"", idStart);
                if (idEnd != std::string::npos)
                {
                    std::string id = relTag.substr(idStart, idEnd - idStart);
                    if (id == rId)
                    {
                        size_t targetStart = relTag.find("Target=\"");
                        if (targetStart != std::string::npos)
                        {
                            targetStart += 8;
                            size_t targetEnd = relTag.find("\"", targetStart);
                            if (targetEnd != std::string::npos)
                            {
                                target = relTag.substr(targetStart, targetEnd - targetStart);
                                break;
                            }
                        }
                    }
                }
            }

            relPos = relEnd + 2;
        }

        if (target.empty())
        {
            std::cerr << "Target not found for rId: " << rId << std::endl;
            return false;
        }

        std::filesystem::path sheetPath = workbook->getDocument().getTempDir() / "xl" / target;

        // Read the current sheet content
        std::ifstream sheetFile(sheetPath);
        if (!sheetFile.is_open())
        {
            std::cerr << "Failed to open sheet: " << sheetPath << std::endl;
            return false;
        }

        std::string sheetContent((std::istreambuf_iterator<char>(sheetFile)), std::istreambuf_iterator<char>());
        sheetFile.close();

        // Find sheetData section
        size_t dataPos = sheetContent.find("<sheetData>");
        if (dataPos == std::string::npos)
        {
            std::cerr << "No sheetData found in sheet: " << name << std::endl;
            return false;
        }

        size_t dataEnd = sheetContent.find("</sheetData>", dataPos);
        if (dataEnd == std::string::npos)
        {
            std::cerr << "No end of sheetData found in sheet: " << name << std::endl;
            return false;
        }

        // Build new sheetData content
        std::string newSheetData = "<sheetData>\n";

        // Process existing cells and add/update our cells
        std::map<std::string, std::string> cellXmlMap;

        // Parse existing cells
        std::string dataSection = sheetContent.substr(dataPos, dataEnd - dataPos + 13);
        size_t cellPos = 0;
        while ((cellPos = dataSection.find("<c ", cellPos)) != std::string::npos)
        {
            size_t tagEnd = dataSection.find("/>", cellPos);
            bool isSelfClosing = (tagEnd != std::string::npos);
            if (!isSelfClosing)
            {
                tagEnd = dataSection.find("</c>", cellPos);
                if (tagEnd == std::string::npos) break;
                tagEnd += 3;
            }
            else
            {
                tagEnd += 1;
            }

            std::string cellTag = dataSection.substr(cellPos, tagEnd - cellPos + 1);

            // Extract r
            size_t rStart = cellTag.find("r=\"");
            std::string ref;
            if (rStart != std::string::npos)
            {
                rStart += 3;
                size_t rEnd = cellTag.find("\"", rStart);
                if (rEnd != std::string::npos)
                {
                    ref = cellTag.substr(rStart, rEnd - rStart);
                }
            }

            if (!ref.empty())
            {
                cellXmlMap[ref] = cellTag;
            }

            cellPos = tagEnd + 1;
        }

        // Update with our cell data
        for (const auto& cellPair : cells)
        {
            const std::string& ref = cellPair.first;
            const XLCell* cell = cellPair.second.get();
            
            // Only save XLCellData cells (not pictures)
            const XLCellData* dataCell = dynamic_cast<const XLCellData*>(cell);
            if (dataCell)
            {
                std::string type = dataCell->getType();
                std::string value = dataCell->getValue();
                
                // Create XML for this cell
                std::string cellXml = "<c r=\"" + ref + "\"";
                if (!type.empty())
                {
                    cellXml += " t=\"" + type + "\"";
                }
                cellXml += ">";
                
                if (!value.empty())
                {
                    cellXml += "<v>" + value + "</v>";
                }
                
                cellXml += "</c>\n";
                
                cellXmlMap[ref] = cellXml;
            }
        }

        // Sort cells by reference for proper ordering
        std::vector<std::pair<std::string, std::string>> sortedCells(cellXmlMap.begin(), cellXmlMap.end());
        std::sort(sortedCells.begin(), sortedCells.end());

        // Build the new sheetData
        for (const auto& cellPair : sortedCells)
        {
            newSheetData += cellPair.second;
        }

        newSheetData += "</sheetData>";

        // Replace the sheetData section
        std::string newSheetContent = sheetContent.substr(0, dataPos) + newSheetData + sheetContent.substr(dataEnd + 13);

        // Write back to file
        std::ofstream outSheetFile(sheetPath);
        if (!outSheetFile.is_open())
        {
            std::cerr << "Failed to write sheet: " << sheetPath << std::endl;
            return false;
        }

        outSheetFile << newSheetContent;
        outSheetFile.close();

        return true;
    }

} // namespace cc::neolux::utils::MiniXLSX