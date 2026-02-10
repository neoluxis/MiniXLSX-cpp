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
#include <map>
#include <algorithm>

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
        // Load shared strings first
        std::filesystem::path sharedStringsPath = workbook->getDocument().getTempDir() / "xl" / "sharedStrings.xml";
        std::ifstream sharedFile(sharedStringsPath);
        if (sharedFile.is_open())
        {
            std::string sharedContent((std::istreambuf_iterator<char>(sharedFile)), std::istreambuf_iterator<char>());
            sharedFile.close();

            size_t siPos = 0;
            while ((siPos = sharedContent.find("<si>", siPos)) != std::string::npos)
            {
                size_t siEnd = sharedContent.find("</si>", siPos);
                if (siEnd == std::string::npos) break;

                std::string siTag = sharedContent.substr(siPos, siEnd - siPos + 5);

                size_t tStart = siTag.find("<t>");
                if (tStart != std::string::npos)
                {
                    tStart += 3;
                    size_t tEnd = siTag.find("</t>", tStart);
                    if (tEnd != std::string::npos)
                    {
                        std::string text = siTag.substr(tStart, tEnd - tStart);
                        sharedStrings.push_back(text);
                    }
                }

                siPos = siEnd + 5;
            }
        }

        // Find the sheet file path using rId
        // In XLSX, xl/_rels/workbook.xml.rels maps rId to target
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

        std::ifstream sheetFile(sheetPath);
        if (!sheetFile.is_open())
        {
            std::cerr << "Failed to open sheet: " << sheetPath << std::endl;
            return false;
        }

        std::string sheetContent((std::istreambuf_iterator<char>(sheetFile)), std::istreambuf_iterator<char>());
        sheetFile.close();

        // Parse sheetData
        size_t dataPos = sheetContent.find("<sheetData>");
        if (dataPos == std::string::npos)
        {
            std::cerr << "No sheetData found." << std::endl;
            return false;
        }

        size_t dataEnd = sheetContent.find("</sheetData>", dataPos);
        if (dataEnd == std::string::npos)
        {
            std::cerr << "No end of sheetData found." << std::endl;
            return false;
        }

        std::string sheetData = sheetContent.substr(dataPos, dataEnd - dataPos + 13);

        size_t cellPos = 0;
        while ((cellPos = sheetData.find("<c ", cellPos)) != std::string::npos)
        {
            // Find end of tag
            size_t tagEnd = sheetData.find("/>", cellPos);
            bool isSelfClosing = (tagEnd != std::string::npos);
            if (!isSelfClosing)
            {
                tagEnd = sheetData.find("</c>", cellPos);
                if (tagEnd == std::string::npos) break;
                tagEnd += 3; // Include </c>
            }
            else
            {
                tagEnd += 1; // Include />
            }

            std::string cellTag = sheetData.substr(cellPos, tagEnd - cellPos + 1);

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

            // Extract t
            std::string type = "n"; // Default to number type
            size_t tPos = cellTag.find(" t=\"");
            if (tPos == std::string::npos) {
                tPos = cellTag.find("t=\"");
                if (tPos == std::string::npos || (tPos > 0 && cellTag[tPos-1] != ' ' && cellTag[tPos-1] != '<')) {
                    // not at start of attribute
                } else {
                    size_t tStart = tPos + 3;
                    size_t tEnd = cellTag.find("\"", tStart);
                    if (tEnd != std::string::npos) {
                        type = cellTag.substr(tStart, tEnd - tStart);
                    }
                }
            } else {
                size_t tStart = tPos + 4;
                size_t tEnd = cellTag.find("\"", tStart);
                if (tEnd != std::string::npos) {
                    type = cellTag.substr(tStart, tEnd - tStart);
                }
            }

            // Extract v
            size_t vStart = cellTag.find("<v>");
            std::string value;
            if (vStart != std::string::npos)
            {
                vStart += 3;
                size_t vEnd = cellTag.find("</v>", vStart);
                if (vEnd != std::string::npos)
                {
                    value = cellTag.substr(vStart, vEnd - vStart);
                }
            }

            if (!ref.empty())
            {
                cells[ref] = std::make_unique<XLCellData>(ref, value, type, sharedStrings);
            }

            cellPos = tagEnd + 1;
        }

        // Parse drawings
        size_t drawingPos = sheetContent.find("<drawing r:id=\"");
        if (drawingPos != std::string::npos)
        {
            size_t idStart = drawingPos + 15; // <drawing r:id="
            size_t idEnd = sheetContent.find("\"", idStart);
            if (idEnd != std::string::npos)
            {
                std::string drawingRId = sheetContent.substr(idStart, idEnd - idStart);

                // Find drawing file using sheet rels
                std::filesystem::path sheetRelsPath = workbook->getDocument().getTempDir() / "xl" / "worksheets" / "_rels" / (std::filesystem::path(target).filename().string() + ".rels");
                std::ifstream sheetRelsFile(sheetRelsPath);
                if (sheetRelsFile.is_open())
                {
                    std::string sheetRelsContent((std::istreambuf_iterator<char>(sheetRelsFile)), std::istreambuf_iterator<char>());
                    sheetRelsFile.close();

                    std::string drawingTarget;
                    size_t relPos = 0;
                    while ((relPos = sheetRelsContent.find("<Relationship ", relPos)) != std::string::npos)
                    {
                        size_t relEnd = sheetRelsContent.find("/>", relPos);
                        if (relEnd == std::string::npos) break;

                        std::string relTag = sheetRelsContent.substr(relPos, relEnd - relPos + 2);

                        size_t idStartRel = relTag.find("Id=\"");
                        if (idStartRel != std::string::npos)
                        {
                            idStartRel += 4;
                            size_t idEndRel = relTag.find("\"", idStartRel);
                            if (idEndRel != std::string::npos)
                            {
                                std::string id = relTag.substr(idStartRel, idEndRel - idStartRel);
                                if (id == drawingRId)
                                {
                                    size_t targetStart = relTag.find("Target=\"");
                                    if (targetStart != std::string::npos)
                                    {
                                        targetStart += 8;
                                        size_t targetEnd = relTag.find("\"", targetStart);
                                        if (targetEnd != std::string::npos)
                                        {
                                            drawingTarget = relTag.substr(targetStart, targetEnd - targetStart);
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        relPos = relEnd + 2;
                    }

                    if (!drawingTarget.empty())
                    {
                        std::filesystem::path drawingPath = (workbook->getDocument().getTempDir() / "xl" / "worksheets" / drawingTarget).lexically_normal();
                        std::filesystem::path drawingRelsPath = drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");

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