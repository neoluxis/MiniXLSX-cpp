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
        std::string tempDir; // Temporary directory for unzipped content
    };

    OpenXLSXWrapper::OpenXLSXWrapper() : impl_(new Impl()) {}
    OpenXLSXWrapper::~OpenXLSXWrapper() { close(); delete impl_; }

    bool OpenXLSXWrapper::ensureTempDir() const
    {
        if (!impl_->tempDir.empty()) return true; // Already exists
        if (!impl_->doc) return false;
        try {
            namespace fs = std::filesystem;
            fs::path tmp = fs::temp_directory_path() / ("minixlsx_unzip_" + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()));
            fs::create_directories(tmp);
            if (cc::neolux::utils::KFZippa::unzip(impl_->openedPath, tmp.string())) {
                impl_->tempDir = tmp.string();
                return true;
            }
        } catch (...) {}
        return false;
    }

    std::string OpenXLSXWrapper::getTempDir() const
    {
        if (!ensureTempDir()) return "";
        return impl_->tempDir;
    }

    void OpenXLSXWrapper::cleanupTempDir()
    {
        if (!impl_->tempDir.empty()) {
            try {
                namespace fs = std::filesystem;
                fs::remove_all(impl_->tempDir);
            } catch (...) {}
            impl_->tempDir.clear();
        }
    }

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
        cleanupTempDir();
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

    std::optional<unsigned int> OpenXLSXWrapper::sheetIndex(const std::string& sheetName) const
    {
        if (!impl_->doc) return std::nullopt;
        try {
            auto names = impl_->doc->workbook().sheetNames();
            for (unsigned int i = 0; i < names.size(); ++i) {
                if (std::string(names[i]) == sheetName) {
                    return i;
                }
            }
        } catch (...) {}
        return std::nullopt;
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

        // Find the drawing file for this sheet
        std::string drawingPath = findDrawingPathForSheet(sheetIndex);
        if (drawingPath.empty()) return out;

        // Parse drawing XML directly
        return parseDrawingXML(drawingPath);
    }

    std::vector<SheetPicture> OpenXLSXWrapper::fetchAllPicturesInSheet(const std::string& sheetName) const
    {
        std::vector<SheetPicture> out;
        if (!impl_->doc) return out;

        // Find sheet index by name
        auto sheetIndexOpt = sheetIndex(sheetName);
        if (!sheetIndexOpt.has_value()) return out;
        unsigned int sheetIndex = sheetIndexOpt.value();

        // Find the drawing file for this sheet
        std::string drawingPath = findDrawingPathForSheet(sheetIndex);
        if (drawingPath.empty()) return out;

        // Parse drawing XML directly
        return parseDrawingXMLForSheetPictures(drawingPath);
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

    std::vector<PictureInfo> OpenXLSXWrapper::parseDrawingXML(const std::string& drawingPath) const
    {
        std::vector<PictureInfo> out;
        if (!ensureTempDir()) return out;

        namespace fs = std::filesystem;
        fs::path tmp(impl_->tempDir);
        fs::path drawingFullPath = tmp / drawingPath;

        if (!fs::exists(drawingFullPath)) return out;

        // Build image map from drawing rels
        fs::path drawingRelsPath = drawingFullPath.parent_path() / "_rels" / (drawingFullPath.filename().string() + ".rels");
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

        // Parse drawing XML
        std::ifstream drawingFile(drawingFullPath);
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
                        absImagePath = (drawingFullPath.parent_path() / imageTarget).lexically_normal();
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
                    PictureInfo pi;
                    pi.ref = ref;
                    pi.fileName = imageFileName;
                    pi.relativePath = relPath;
                    out.push_back(pi);
                }
            }

            anchorPos = anchorEnd + 20;
        }

        return out;
    }

    std::string OpenXLSXWrapper::findDrawingPathForSheet(unsigned int sheetIndex) const
    {
        if (!ensureTempDir()) return "";

        namespace fs = std::filesystem;
        fs::path tmp(impl_->tempDir);
        // We will NOT read the sheet XML relationships to find the drawing.
        // Instead assume sheet index N -> sheetN.xml and drawingN.xml under xl/drawings.
        // Construct expected drawing path directly: xl/drawings/drawing{N}.xml
        std::string drawingFile = "drawings/drawing" + std::to_string(sheetIndex + 1) + ".xml";
        fs::path drawingPath = tmp / "xl" / drawingFile;
        if (!fs::exists(drawingPath)) {
            // fallback: some files may use different numbering; also check drawings/drawing1.xml if index out-of-range
            if (sheetIndex == 0) {
                fs::path alt = tmp / "xl" / "drawings" / "drawing1.xml";
                if (fs::exists(alt)) drawingPath = alt;
            }
        }

        if (!fs::exists(drawingPath)) return "";

        return drawingPath.string().substr(tmp.string().size() + 1); // Return relative path from temp dir
    }

    std::vector<SheetPicture> OpenXLSXWrapper::parseDrawingXMLForSheetPictures(const std::string& drawingPath) const
    {
        std::vector<SheetPicture> out;
        if (!ensureTempDir()) return out;

        namespace fs = std::filesystem;
        fs::path tmp(impl_->tempDir);
        fs::path drawingFullPath = tmp / drawingPath;

        if (!fs::exists(drawingFullPath)) return out;

        // Build image map from drawing rels
        fs::path drawingRelsPath = drawingFullPath.parent_path() / "_rels" / (drawingFullPath.filename().string() + ".rels");
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

        // Parse drawing XML
        std::ifstream drawingFile(drawingFullPath);
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
                        absImagePath = (drawingFullPath.parent_path() / imageTarget).lexically_normal();
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

        return out;
    }
} // namespace cc::neolux::utils::MiniXLSX
