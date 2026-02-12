#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include <memory>
#include <iostream>
#include <unordered_map>

// Use OpenXLSX when available
#include "OpenXLSX.hpp"
#include <filesystem>
#include <chrono>
#include <pugixml.hpp>
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include <fstream>
#include "cc/neolux/utils/MiniXLSX/XLSheet.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    struct OpenXLSXWrapper::Impl {
        std::unique_ptr<OpenXLSX::XLDocument> doc;
        std::string openedPath;
    };

    OpenXLSXWrapper::OpenXLSXWrapper() : impl_(new Impl()) {}
    OpenXLSXWrapper::~OpenXLSXWrapper() { close(); delete impl_; }

    bool OpenXLSXWrapper::open(const std::string& path)
    {
        try {
            impl_->doc = std::make_unique<OpenXLSX::XLDocument>();
            impl_->doc->open(path);
            impl_->openedPath = path;
            return static_cast<bool>(impl_->doc && impl_->doc->isOpen());
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::open error: " << e.what() << std::endl;
            impl_->doc.reset();
            return false;
        }
    }

    void OpenXLSXWrapper::close()
    {
        if (impl_->doc)
        {
            try { impl_->doc->close(); } catch (...) {}
            impl_->doc.reset();
        }
    }

    bool OpenXLSXWrapper::isOpen() const
    {
        return impl_->doc && impl_->doc->isOpen();
    }

    unsigned int OpenXLSXWrapper::sheetCount() const
    {
        if (!impl_->doc) return 0;
        try {
            return impl_->doc->workbook().sheetCount();
        } catch (...) { return 0; }
    }

    std::string OpenXLSXWrapper::sheetName(unsigned int index) const
    {
        try {
            if (!impl_->doc) return std::string();
            auto names = impl_->doc->workbook().sheetNames();
            if (index < names.size()) return std::string(names[index]);
        } catch (...) {}
        return std::string();
    }

    std::optional<std::string> OpenXLSXWrapper::getCellValue(unsigned int sheetIndex, const std::string& ref) const
    {
        if (!impl_->doc) return std::nullopt;
        try {
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            auto cell = ws.cell(ref);
            std::string val = cell.getString();
            return std::optional<std::string>(val);
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::getCellValue error: " << e.what() << std::endl;
            return std::nullopt;
        }
    }

    bool OpenXLSXWrapper::setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value)
    {
        if (!impl_->doc) return false;
        try {
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            ws.cell(ref) = value;
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::setCellValue error: " << e.what() << std::endl;
            return false;
        }
    }

    bool OpenXLSXWrapper::setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style)
    {
        if (!impl_->doc) return false;
        try {
            auto &styles = impl_->doc->styles();

            // Create fill
            OpenXLSX::XLFills &fills = styles.fills();
            OpenXLSX::XLStyleIndex fillIdx = fills.create();
            if (!style.backgroundColor.empty()) {
                std::string col = style.backgroundColor;
                if (!col.empty() && col[0] == '#') col = col.substr(1);
                if (col.size() == 6) col = std::string("FF") + col;
                fills[fillIdx].setPatternType(OpenXLSX::XLPatternSolid);
                fills[fillIdx].setColor(OpenXLSX::XLColor(col));
            }

            // Create border
            OpenXLSX::XLBorders &borders = styles.borders();
            OpenXLSX::XLStyleIndex borderIdx = borders.create();
            bool hasBorder = (style.border != CellBorderStyle::None);
            if (hasBorder) {
                OpenXLSX::XLLineStyle ls = OpenXLSX::XLLineStyleThin;
                switch (style.border) {
                    case CellBorderStyle::Thin: ls = OpenXLSX::XLLineStyleThin; break;
                    case CellBorderStyle::Medium: ls = OpenXLSX::XLLineStyleMedium; break;
                    case CellBorderStyle::Thick: ls = OpenXLSX::XLLineStyleThick; break;
                    default: ls = OpenXLSX::XLLineStyleThin; break;
                }
                std::string bcol = style.borderColor.empty() ? std::string("FF000000") : style.borderColor;
                if (!bcol.empty() && bcol[0] == '#') bcol = bcol.substr(1);
                if (bcol.size() == 6) bcol = std::string("FF") + bcol;
                OpenXLSX::XLColor bcolc(bcol);
                borders[borderIdx].setLeft(ls, bcolc);
                borders[borderIdx].setRight(ls, bcolc);
                borders[borderIdx].setTop(ls, bcolc);
                borders[borderIdx].setBottom(ls, bcolc);
            }

            // Create cell format — if cell already has a format, copy it to preserve font and other properties
            OpenXLSX::XLCellFormats &cellFormats = styles.cellFormats();
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            OpenXLSX::XLStyleIndex baseFmt = 0;
            try {
                baseFmt = ws.cell(ref).cellFormat();
            } catch(...) { baseFmt = 0; }

            OpenXLSX::XLStyleIndex xf;
            try {
                if (baseFmt > 0) {
                    xf = cellFormats.create(cellFormats[baseFmt]);
                } else {
                    xf = cellFormats.create();
                }
            } catch(...) {
                xf = cellFormats.create();
            }

            if (!style.backgroundColor.empty()) cellFormats[xf].setFillIndex(fillIdx);
            if (hasBorder) cellFormats[xf].setBorderIndex(borderIdx);
            cellFormats[xf].setApplyFill(!style.backgroundColor.empty());
            if (hasBorder) cellFormats[xf].setApplyBorder(true);

            // Ensure font is applied (preserve existing font)
            try {
                cellFormats[xf].setApplyFont(true);
            } catch(...) {}

            ws.cell(ref).setCellFormat(xf);
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::setCellStyle error: " << e.what() << std::endl;
            return false;
        }
    }

    bool OpenXLSXWrapper::save()
    {
        if (!impl_->doc) return false;
        try {
            impl_->doc->save();
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::save error: " << e.what() << std::endl;
            return false;
        }
    }

    std::vector<PictureInfo> OpenXLSXWrapper::getPictures(unsigned int sheetIndex) const
    {
        std::vector<PictureInfo> out;
        if (!impl_->doc) return out;
        try {
            namespace fs = std::filesystem;
            // Unzip to temp dir and parse drawing XML similarly to XLSheet::load
            fs::path tmp = fs::temp_directory_path() / ("minixlsx_unzip_" + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()));
            try { fs::create_directories(tmp); } catch(...) {}
            if (!cc::neolux::utils::KFZippa::unzip(impl_->openedPath, tmp.string())) {
                // failed to extract, fallback to content types scan
                OpenXLSX::XLZipArchive za;
                za.open(impl_->openedPath);
                std::string ctRaw = za.getEntry("[Content_Types].xml");
                if (ctRaw.empty()) return out;
                OpenXLSX::XLXmlData tmpXml(nullptr, "[Content_Types].xml");
                tmpXml.setRawData(ctRaw);
                OpenXLSX::XLContentTypes contentTypes(&tmpXml);
                auto items = contentTypes.getContentItems();
                for (const auto &item : items) {
                    std::string path = item.path();
                    if (path.rfind("xl/media/", 0) == 0) {
                        PictureInfo pi;
                        pi.relativePath = path;
                        auto pos = path.find_last_of('/');
                        if (pos != std::string::npos) pi.fileName = path.substr(pos + 1);
                        else pi.fileName = path;
                        out.push_back(pi);
                    }
                }
                return out;
            } else {
                // Attempt to load worksheet file by index (sheetN.xml)
                std::string sheetFile = "sheet" + std::to_string(sheetIndex + 1) + ".xml";
                fs::path sheetPath = tmp / "xl" / "worksheets" / sheetFile;
                if (!fs::exists(sheetPath)) {
                    // try workbook rels to map sheet index to target
                    fs::path relsPath = tmp / "xl" / "_rels" / "workbook.xml.rels";
                    if (fs::exists(relsPath)) {
                        pugi::xml_document relsDoc;
                        if (relsDoc.load_file(relsPath.c_str())) {
                            pugi::xml_node root = relsDoc.child("Relationships");
                            unsigned int idx = 0;
                            for (pugi::xml_node rel : root.children("Relationship")) {
                                if (rel.attribute("Type").as_string() == std::string("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")) {
                                    if (idx == sheetIndex) {
                                        std::string target = rel.attribute("Target").as_string();
                                        sheetPath = (tmp / "xl" / target).lexically_normal();
                                        break;
                                    }
                                    ++idx;
                                }
                            }
                        }
                    }
                }

                if (fs::exists(sheetPath)) {
                    pugi::xml_document sheetDoc;
                    if (sheetDoc.load_file(sheetPath.c_str())) {
                        pugi::xml_node worksheet = sheetDoc.child("worksheet");
                        pugi::xml_node drawingNode = worksheet.child("drawing");
                        if (drawingNode) {
                            std::string drawingRId = drawingNode.attribute("r:id").as_string();
                            // find drawing target in sheet rels
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
                                fs::path drawingPath;
                                try {
                                    drawingPath = (sheetPath.parent_path() / drawingTarget).lexically_normal();
                                } catch(...) {
                                    drawingPath = (tmp / "xl" / drawingTarget).lexically_normal();
                                }
                                fs::path drawingRelsPath = drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");

                                // Build image map from drawing rels
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
                                    while ((anchorPos = drawingContent.find("<xdr:twoCellAnchor", anchorPos)) != std::string::npos) {
                                        size_t anchorEnd = drawingContent.find("</xdr:twoCellAnchor>", anchorPos);
                                        if (anchorEnd == std::string::npos) break;
                                        std::string anchorTag = drawingContent.substr(anchorPos, anchorEnd - anchorPos + 20);

                                        // Extract from col and row
                                        size_t fromPos = anchorTag.find("<xdr:from>");
                                        int fromCol = -1, fromRow = -1;
                                        if (fromPos != std::string::npos) {
                                            size_t colPos = anchorTag.find("<xdr:col>", fromPos);
                                            if (colPos != std::string::npos) {
                                                colPos += 9;
                                                size_t colEnd = anchorTag.find("</xdr:col>", colPos);
                                                if (colEnd != std::string::npos) {
                                                    fromCol = std::stoi(anchorTag.substr(colPos, colEnd - colPos));
                                                }
                                            }
                                            size_t rowPos = anchorTag.find("<xdr:row>", fromPos);
                                            if (rowPos != std::string::npos) {
                                                rowPos += 9;
                                                size_t rowEnd = anchorTag.find("</xdr:row>", rowPos);
                                                if (rowEnd != std::string::npos) {
                                                    fromRow = std::stoi(anchorTag.substr(rowPos, rowEnd - rowPos));
                                                }
                                            }
                                        }

                                        // Extract embed id
                                        size_t embedPos = anchorTag.find("r:embed=\"");
                                        std::string embedId;
                                        if (embedPos != std::string::npos) {
                                            embedPos += 9;
                                            size_t embedEnd = anchorTag.find("\"", embedPos);
                                            if (embedEnd != std::string::npos) embedId = anchorTag.substr(embedPos, embedEnd - embedPos);
                                        }

                                        if (fromCol >= 0 && fromRow >= 0 && !embedId.empty()) {
                                            std::string ref = XLSheet::columnNumberToLetter(fromCol) + std::to_string(fromRow + 1);
                                            auto it = imageMap.find(embedId);
                                            if (it != imageMap.end()) {
                                                std::string imageTarget = it->second; // e.g., "../media/image1.jpg"
                                                fs::path absImagePath;
                                                try {
                                                    absImagePath = (drawingPath.parent_path() / imageTarget).lexically_normal();
                                                } catch(...) {
                                                    absImagePath = fs::path("xl") / imageTarget;
                                                }
                                                std::string imageFileName = absImagePath.filename().string();
                                                std::string relPath;
                                                try {
                                                    fs::path base = tmp / "xl";
                                                    fs::path parent = absImagePath.parent_path();
                                                    relPath = fs::relative(parent, base).string();
                                                } catch(...) {
                                                    relPath = absImagePath.parent_path().string();
                                                    if (relPath == "../media") relPath = "media";
                                                }
                                                PictureInfo pi;
                                                pi.ref = ref;
                                                pi.fileName = imageFileName;
                                                pi.relativePath = relPath;
                                                out.push_back(pi);
                                            }
                                        }

                                        anchorPos = anchorEnd + 20;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // cleanup
            try { std::filesystem::remove_all(tmp); } catch(...) {}
        } catch (...) {}
        return out;
    }

    std::vector<SheetPicture> OpenXLSXWrapper::fetchAllPicturesInSheet(const std::string& sheetName) const
    {
        std::vector<SheetPicture> out;
        if (!impl_->doc) return out;

        // Find sheet index by name
        int sheetIndex = -1;
        for (unsigned int i = 0; i < sheetCount(); ++i) {
            if (this->sheetName(i) == sheetName) {
                sheetIndex = static_cast<int>(i);
                break;
            }
        }
        if (sheetIndex < 0) return out;

        // Get pictures using existing logic, but parse to extract row/col details
        try {
            namespace fs = std::filesystem;
            // Unzip to temp dir and parse drawing XML similarly to XLSheet::load
            fs::path tmp = fs::temp_directory_path() / ("minixlsx_unzip_" + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()));
            try { fs::create_directories(tmp); } catch(...) {}
            if (!cc::neolux::utils::KFZippa::unzip(impl_->openedPath, tmp.string())) {
                // If unzip fails, cannot parse XML, return empty
                return out;
            }

            // Attempt to load worksheet file by index (sheetN.xml)
            std::string sheetFile = "sheet" + std::to_string(sheetIndex + 1) + ".xml";
            fs::path sheetPath = tmp / "xl" / "worksheets" / sheetFile;
            if (!fs::exists(sheetPath)) {
                // try workbook rels to map sheet index to target
                fs::path relsPath = tmp / "xl" / "_rels" / "workbook.xml.rels";
                if (fs::exists(relsPath)) {
                    pugi::xml_document relsDoc;
                    if (relsDoc.load_file(relsPath.c_str())) {
                        pugi::xml_node root = relsDoc.child("Relationships");
                        unsigned int idx = 0;
                        for (pugi::xml_node rel : root.children("Relationship")) {
                            if (rel.attribute("Type").as_string() == std::string("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")) {
                                if (idx == sheetIndex) {
                                    std::string target = rel.attribute("Target").as_string();
                                    sheetPath = (tmp / "xl" / target).lexically_normal();
                                    break;
                                }
                                ++idx;
                            }
                        }
                    }
                }
            }

            if (fs::exists(sheetPath)) {
                pugi::xml_document sheetDoc;
                if (sheetDoc.load_file(sheetPath.c_str())) {
                    pugi::xml_node worksheet = sheetDoc.child("worksheet");
                    pugi::xml_node drawingNode = worksheet.child("drawing");
                    if (drawingNode) {
                        std::string drawingRId = drawingNode.attribute("r:id").as_string();
                        // find drawing target in sheet rels
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
                            fs::path drawingPath;
                            try {
                                drawingPath = (sheetPath.parent_path() / drawingTarget).lexically_normal();
                            } catch(...) {
                                drawingPath = (tmp / "xl" / drawingTarget).lexically_normal();
                            }
                            fs::path drawingRelsPath = drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");

                            // Build image map from drawing rels
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
                                while ((anchorPos = drawingContent.find("<xdr:twoCellAnchor", anchorPos)) != std::string::npos) {
                                    size_t anchorEnd = drawingContent.find("</xdr:twoCellAnchor>", anchorPos);
                                    if (anchorEnd == std::string::npos) break;
                                    std::string anchorTag = drawingContent.substr(anchorPos, anchorEnd - anchorPos + 20);

                                    // Extract from col and row
                                    size_t fromPos = anchorTag.find("<xdr:from>");
                                    int fromCol = -1, fromRow = -1;
                                    if (fromPos != std::string::npos) {
                                        size_t colPos = anchorTag.find("<xdr:col>", fromPos);
                                        if (colPos != std::string::npos) {
                                            colPos += 9;
                                            size_t colEnd = anchorTag.find("</xdr:col>", colPos);
                                            if (colEnd != std::string::npos) {
                                                fromCol = std::stoi(anchorTag.substr(colPos, colEnd - colPos));
                                            }
                                        }
                                        size_t rowPos = anchorTag.find("<xdr:row>", fromPos);
                                        if (rowPos != std::string::npos) {
                                            rowPos += 9;
                                            size_t rowEnd = anchorTag.find("</xdr:row>", rowPos);
                                            if (rowEnd != std::string::npos) {
                                                fromRow = std::stoi(anchorTag.substr(rowPos, rowEnd - rowPos));
                                            }
                                        }
                                    }

                                    // Extract embed id
                                    size_t embedPos = anchorTag.find("r:embed=\"");
                                    std::string embedId;
                                    if (embedPos != std::string::npos) {
                                        embedPos += 9;
                                        size_t embedEnd = anchorTag.find("\"", embedPos);
                                        if (embedEnd != std::string::npos) embedId = anchorTag.substr(embedPos, embedEnd - embedPos);
                                    }

                                    if (fromCol >= 0 && fromRow >= 0 && !embedId.empty()) {
                                        auto it = imageMap.find(embedId);
                                        if (it != imageMap.end()) {
                                            std::string imageTarget = it->second; // e.g., "../media/image1.jpg"
                                            fs::path absImagePath;
                                            try {
                                                absImagePath = (drawingPath.parent_path() / imageTarget).lexically_normal();
                                            } catch(...) {
                                                absImagePath = fs::path("xl") / imageTarget;
                                            }
                                            std::string imageFileName = absImagePath.filename().string();
                                            std::string relPath;
                                            try {
                                                fs::path base = tmp / "xl";
                                                fs::path parent = absImagePath.parent_path();
                                                relPath = fs::relative(parent, base).string();
                                            } catch(...) {
                                                relPath = absImagePath.parent_path().string();
                                                if (relPath == "../media") relPath = "media";
                                            }
                                            if (relPath == "../media") relPath = "media";
                                            SheetPicture sp;
                                            sp.row = std::to_string(fromRow + 1);
                                            sp.col = XLSheet::columnNumberToLetter(fromCol);
                                            sp.rowNum = fromRow + 1;
                                            sp.colNum = fromCol + 1;  // Since columns are 1-based in Excel
                                            sp.relativePath = relPath + "/" + imageFileName;
                                            out.push_back(sp);
                                        }
                                    }

                                    anchorPos = anchorEnd + 20;
                                }
                            }
                        }
                    }
                }
            }

            // cleanup
            try { std::filesystem::remove_all(tmp); } catch(...) {}
        } catch (...) {}
        return out;
    }

    std::optional<std::vector<uint8_t>> OpenXLSXWrapper::getPictureRaw(unsigned int sheetIndex, const std::string& ref) const
    {
        if (!impl_->doc) return std::nullopt;
        try {
            auto pics = getPictures(sheetIndex);
            for (const auto &pi : pics) {
                if (pi.ref == ref) {
                    std::string rel = pi.relativePath;
                    if (rel.rfind("..", 0) == 0) {
                        size_t pos = rel.find("/media");
                        if (pos != std::string::npos) rel = std::string("media");
                    }
                    if (rel.empty()) rel = "media";
                    std::string entry = std::string("xl/") + rel + "/" + pi.fileName;
                    OpenXLSX::XLZipArchive za;
                    za.open(impl_->openedPath);
                    std::string data = za.getEntry(entry);
                    if (data.empty()) {
                        // fallback to xl/media/<file>
                        entry = std::string("xl/media/") + pi.fileName;
                        data = za.getEntry(entry);
                    }
                    if (!data.empty()) {
                        std::vector<uint8_t> out(data.begin(), data.end());
                        return out;
                    }
                    return std::nullopt;
                }
            }
        } catch (...) {}
        return std::nullopt;
    }

} // namespace cc::neolux::utils::MiniXLSX
