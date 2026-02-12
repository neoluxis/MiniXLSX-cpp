#include "cc/neolux/utils/MiniXLSX/XLSheet.hpp"
#include "cc/neolux/utils/MiniXLSX/XLWorkbook.hpp"
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellData.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>
#include <memory>
#include <map>
#include <unordered_map>
#include <algorithm>
#include <pugixml.hpp>
#include "cc/neolux/utils/MiniXLSX/Types.hpp"

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
        : workbook(&wb), name(n), sheetId(sid), rId(rid), oxWrapper(nullptr), oxSheetIndex(0)
    {
    }

    XLSheet::XLSheet(XLWorkbook& wb, OpenXLSXWrapper* wrapper, XLPictureReader* picReader, unsigned int sheetIndex)
        : workbook(&wb), name(wrapper ? wrapper->sheetName(sheetIndex) : std::string()), sheetId(), rId(), oxWrapper(wrapper), pictureReader(picReader), oxSheetIndex(sheetIndex)
    {
    }

    XLSheet::~XLSheet()
    {
    }

    bool XLSheet::load()
    {
        // 若由 OpenXLSXWrapper 提供数据，则无需从 XML 加载
        if (oxWrapper && oxWrapper->isOpen()) {
            return true;
        }
        namespace fs = std::filesystem;
        fs::path temp = workbook->getDocument().getTempDir();

        // 使用 pugixml 加载共享字符串
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
                    // 优先读取 <t>，并兼容富文本
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

        // 通过 rId 在 workbook 关系中定位工作表路径
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

        // 解析 sheetData
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

        // 解析图片（若存在），使用关系文件映射图片目标
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

                        // 读取 drawing 的关系文件，构建图片映射
                        std::unordered_map<std::string, std::string> imageMap;
                        if (fs::exists(drawingRelsPath)) {
                            pugi::xml_document drelDoc;
                            if (drelDoc.load_file(drawingRelsPath.c_str())) {
                                pugi::xml_node droot = drelDoc.child("Relationships");
                                for (pugi::xml_node rel : droot.children("Relationship")) {
                                    std::string type = rel.attribute("Type").as_string();
                                    if (type.find("image") != std::string::npos) {
                                        imageMap[rel.attribute("Id").as_string()] = rel.attribute("Target").as_string();
                                    }
                                }
                            }
                        }

                        // 读取 drawing XML，查找 <xdr:twoCellAnchor>
                        if (fs::exists(drawingPath)) {
                            std::ifstream drawingFile(drawingPath);
                            std::string drawingContent((std::istreambuf_iterator<char>(drawingFile)), std::istreambuf_iterator<char>());
                            drawingFile.close();

                            size_t anchorPos = 0;
                            while ((anchorPos = drawingContent.find("<xdr:twoCellAnchor", anchorPos)) != std::string::npos)
                            {
                                size_t anchorEnd = drawingContent.find("</xdr:twoCellAnchor>", anchorPos);
                                if (anchorEnd == std::string::npos) break;

                                std::string anchorTag = drawingContent.substr(anchorPos, anchorEnd - anchorPos + 20);

                                // 提取起始列与行
                                size_t fromPos = anchorTag.find("<xdr:from>");
                                int fromCol = -1, fromRow = -1;
                                if (fromPos != std::string::npos)
                                {
                                    size_t colPos = anchorTag.find("<xdr:col>", fromPos);
                                    if (colPos != std::string::npos)
                                    {
                                        colPos += 9;
                                        size_t colEnd = anchorTag.find("</xdr:col>", colPos);
                                        if (colEnd != std::string::npos)
                                        {
                                            fromCol = std::stoi(anchorTag.substr(colPos, colEnd - colPos));
                                        }
                                    }
                                    size_t rowPos = anchorTag.find("<xdr:row>", fromPos);
                                    if (rowPos != std::string::npos)
                                    {
                                        rowPos += 9;
                                        size_t rowEnd = anchorTag.find("</xdr:row>", rowPos);
                                        if (rowEnd != std::string::npos)
                                        {
                                            fromRow = std::stoi(anchorTag.substr(rowPos, rowEnd - rowPos));
                                        }
                                    }
                                }

                                // 提取嵌入关系 ID
                                size_t embedPos = anchorTag.find("r:embed=\"");
                                std::string embedId;
                                if (embedPos != std::string::npos)
                                {
                                    embedPos += 9;
                                    size_t embedEnd = anchorTag.find("\"", embedPos);
                                    if (embedEnd != std::string::npos)
                                    {
                                        embedId = anchorTag.substr(embedPos, embedEnd - embedPos);
                                    }
                                }

                                if (fromCol >= 0 && fromRow >= 0 && !embedId.empty())
                                {
                                    std::string ref = XLSheet::columnNumberToLetter(fromCol) + std::to_string(fromRow + 1);
                                    auto it = imageMap.find(embedId);
                                    if (it != imageMap.end())
                                    {
                                        std::string imageTarget = it->second; // 例如 "../media/image1.jpg"
                                        std::filesystem::path imagePath(imageTarget);
                                        std::string imageFileName = imagePath.filename().string();
                                        std::string relPath = imagePath.parent_path().string();
                                        if (relPath == "../media") relPath = "media";
                                        PictureInfo pi;
                                        pi.ref = ref;
                                        pi.fileName = imageFileName;
                                        pi.relativePath = relPath;
                                        cells[ref] = std::make_unique<XLCellPicture>(pi.ref, pi.fileName, pi.relativePath);
                                    }
                                }

                                anchorPos = anchorEnd + 20;
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
        // 若由封装提供数据，则按需获取并缓存
        if (oxWrapper && oxWrapper->isOpen()) {
            // 仅在首次访问时加载图片缓存
            if (!picturesLoaded) {
                try {
                    if (pictureReader) {
                        auto pics = pictureReader->getPictures(oxSheetIndex);
                        auto nonConstThis = const_cast<XLSheet*>(this);
                        for (const auto &pi : pics) {
                            nonConstThis->cells[pi.ref] = std::make_unique<XLCellPicture>(pi.ref, pi.fileName, pi.relativePath);
                        }
                    }
                } catch (...) {}
                picturesLoaded = true;
            }

            auto it = cells.find(ref);
            if (it != cells.end()) return it->second.get();
            auto v = oxWrapper->getCellValue(oxSheetIndex, ref);
            if (v.has_value()) {
                // 在 const 方法中创建可变缓存
                auto nonConstThis = const_cast<XLSheet*>(this);
                nonConstThis->cells[ref] = std::make_unique<XLCellData>(ref, v.value(), std::string("str"), std::vector<std::string>());
                return nonConstThis->cells[ref].get();
            }
            return nullptr;
        }

        // 兜底：若封装未提供图片信息，尝试从解压目录解析
        try {
            namespace fs = std::filesystem;
            fs::path temp = workbook->getDocument().getTempDir();
            fs::path sheetPath = temp / "xl" / "worksheets" / (std::string("sheet") + std::to_string(oxSheetIndex + 1) + ".xml");
            if (fs::exists(sheetPath)) {
                pugi::xml_document sheetDoc;
                if (sheetDoc.load_file(sheetPath.c_str())) {
                    pugi::xml_node worksheet = sheetDoc.child("worksheet");
                    pugi::xml_node drawingNode = worksheet.child("drawing");
                    if (drawingNode) {
                        std::string drawingRId = drawingNode.attribute("r:id").as_string();
                        fs::path sheetRelsPath = sheetPath.parent_path() / "_rels" / (sheetPath.filename().string() + ".rels");
                        std::string drawingTarget;
                        if (fs::exists(sheetRelsPath)) {
                            pugi::xml_document sheetRelsDoc;
                            if (sheetRelsDoc.load_file(sheetRelsPath.c_str())) {
                                pugi::xml_node sroot = sheetRelsDoc.child("Relationships");
                                for (pugi::xml_node rel : sroot.children("Relationship")) {
                                    if (std::string(rel.attribute("Id").as_string()) == drawingRId) {
                                        drawingTarget = rel.attribute("Target").as_string();
                                        break;
                                    }
                                }
                            }
                        }

                        if (!drawingTarget.empty()) {
                            fs::path drawingPath = (temp / "xl" / drawingTarget).lexically_normal();
                            fs::path drawingRelsPath = drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");
                            std::unordered_map<std::string, std::string> imageMap;
                            if (fs::exists(drawingRelsPath)) {
                                pugi::xml_document drelDoc;
                                if (drelDoc.load_file(drawingRelsPath.c_str())) {
                                    pugi::xml_node droot = drelDoc.child("Relationships");
                                    for (pugi::xml_node rel : droot.children("Relationship")) {
                                        std::string type = rel.attribute("Type").as_string();
                                        if (type.find("image") != std::string::npos) {
                                            imageMap[rel.attribute("Id").as_string()] = rel.attribute("Target").as_string();
                                        }
                                    }
                                }
                            }

                            if (fs::exists(drawingPath)) {
                                std::ifstream drawingFile(drawingPath);
                                std::string drawingContent((std::istreambuf_iterator<char>(drawingFile)), std::istreambuf_iterator<char>());
                                drawingFile.close();
                                size_t anchorPos = 0;
                                auto nonConstThis = const_cast<XLSheet*>(this);
                                while ((anchorPos = drawingContent.find("<xdr:twoCellAnchor", anchorPos)) != std::string::npos) {
                                    size_t anchorEnd = drawingContent.find("</xdr:twoCellAnchor>", anchorPos);
                                    if (anchorEnd == std::string::npos) break;
                                    std::string anchorTag = drawingContent.substr(anchorPos, anchorEnd - anchorPos + 20);
                                    size_t fromPos = anchorTag.find("<xdr:from>");
                                    int fromCol = -1, fromRow = -1;
                                    if (fromPos != std::string::npos) {
                                        size_t colPos = anchorTag.find("<xdr:col>", fromPos);
                                        if (colPos != std::string::npos) {
                                            colPos += 9;
                                            size_t colEnd = anchorTag.find("</xdr:col>", colPos);
                                            if (colEnd != std::string::npos) fromCol = std::stoi(anchorTag.substr(colPos, colEnd - colPos));
                                        }
                                        size_t rowPos = anchorTag.find("<xdr:row>", fromPos);
                                        if (rowPos != std::string::npos) {
                                            rowPos += 9;
                                            size_t rowEnd = anchorTag.find("</xdr:row>", rowPos);
                                            if (rowEnd != std::string::npos) fromRow = std::stoi(anchorTag.substr(rowPos, rowEnd - rowPos));
                                        }
                                    }
                                    size_t embedPos = anchorTag.find("r:embed=\"");
                                    std::string embedId;
                                    if (embedPos != std::string::npos) {
                                        embedPos += 9;
                                        size_t embedEnd = anchorTag.find("\"", embedPos);
                                        if (embedEnd != std::string::npos) embedId = anchorTag.substr(embedPos, embedEnd - embedPos);
                                    }
                                    if (fromCol >= 0 && fromRow >= 0 && !embedId.empty()) {
                                        std::string cellRef = XLSheet::columnNumberToLetter(fromCol) + std::to_string(fromRow + 1);
                                        auto itImg = imageMap.find(embedId);
                                        if (itImg != imageMap.end()) {
                                            std::filesystem::path imagePath(itImg->second);
                                            std::string imageFileName = imagePath.filename().string();
                                            std::string relPath = imagePath.parent_path().string();
                                            if (relPath == "../media") relPath = "media";
                                            nonConstThis->cells[cellRef] = std::make_unique<XLCellPicture>(cellRef, imageFileName, relPath);
                                        }
                                    }
                                    anchorPos = anchorEnd + 20;
                                }
                            }
                        }
                    }
                }
            }
        } catch(...) {}

        return nullptr;

        auto it = cells.find(ref);
        if (it != cells.end())
        {
            return it->second.get();
        }
        return nullptr;
    }

    std::string XLSheet::getCellValue(const std::string& ref) const
    {
        if (oxWrapper && oxWrapper->isOpen()) {
            auto v = oxWrapper->getCellValue(oxSheetIndex, ref);
            if (v.has_value()) return v.value();
            return std::string();
        }

        const XLCell* cell = getCell(ref);
        if (cell)
        {
            return cell->getValue();
        }
        return "";
    }

    void XLSheet::setCellValue(const std::string& ref, const std::string& value, const std::string& type)
    {
        if (oxWrapper && oxWrapper->isOpen()) {
            // 委托给封装处理
            oxWrapper->setCellValue(oxSheetIndex, ref, value);
            workbook->getDocument().markModified();
            return;
        }

        auto it = cells.find(ref);
        if (it != cells.end())
        {
            // 单元格存在则更新
            XLCellData* dataCell = dynamic_cast<XLCellData*>(it->second.get());
            if (dataCell)
            {
                dataCell->setValue(value);
                dataCell->setType(type);
            }
        }
        else
        {
            // 单元格不存在则创建
            cells[ref] = std::make_unique<XLCellData>(ref, value, type, sharedStrings);
        }
        
        // 标记文档已修改
        workbook->getDocument().markModified();
    }

    bool XLSheet::save()
    {
        // 通过 rId 定位工作表文件路径
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

        // 读取当前工作表内容
        std::ifstream sheetFile(sheetPath);
        if (!sheetFile.is_open())
        {
            std::cerr << "Failed to open sheet: " << sheetPath << std::endl;
            return false;
        }

        std::string sheetContent((std::istreambuf_iterator<char>(sheetFile)), std::istreambuf_iterator<char>());
        sheetFile.close();

        // 查找 sheetData 段
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

        // 构建新的 sheetData 内容
        std::string newSheetData = "<sheetData>\n";

        // 处理已有单元格并合并更新
        std::map<std::string, std::string> cellXmlMap;

        // 解析已有单元格
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

            // 提取 r
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

        // 合并当前内存中的单元格数据
        for (const auto& cellPair : cells)
        {
            const std::string& ref = cellPair.first;
            const XLCell* cell = cellPair.second.get();
            
            // 仅保存 XLCellData（不保存图片单元格）
            const XLCellData* dataCell = dynamic_cast<const XLCellData*>(cell);
            if (dataCell)
            {
                std::string type = dataCell->getType();
                std::string value = dataCell->getValue();
                
                // 生成单元格 XML
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

        // 按引用排序
        std::vector<std::pair<std::string, std::string>> sortedCells(cellXmlMap.begin(), cellXmlMap.end());
        std::sort(sortedCells.begin(), sortedCells.end());

        // 拼接新的 sheetData
        for (const auto& cellPair : sortedCells)
        {
            newSheetData += cellPair.second;
        }

        newSheetData += "</sheetData>";

        // 替换 sheetData 段
        std::string newSheetContent = sheetContent.substr(0, dataPos) + newSheetData + sheetContent.substr(dataEnd + 13);

        // 写回文件
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