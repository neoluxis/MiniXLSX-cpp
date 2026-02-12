#include "cc/neolux/utils/MiniXLSX/XLPictureReader.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include "cc/neolux/utils/MiniXLSX/XLSheet.hpp"
#include "OpenXLSX.hpp"
#include <filesystem>
#include <chrono>
#include <fstream>
#include <unordered_map>
#include <pugixml.hpp>

namespace cc::neolux::utils::MiniXLSX
{
    XLPictureReader::XLPictureReader() = default;
    XLPictureReader::~XLPictureReader() { close(); }

    bool XLPictureReader::open(const std::string& xlsxPath)
    {
        close();
        openedPath = xlsxPath;
        attached = false;
        ownsTempDir = true;
        return true;
    }

    bool XLPictureReader::attach(const std::string& xlsxPath, const std::string& temp)
    {
        close();
        openedPath = xlsxPath;
        tempDir = temp;
        attached = true;
        ownsTempDir = false;
        return true;
    }

    void XLPictureReader::close()
    {
        if (ownsTempDir) cleanupTempDir();
        openedPath.clear();
        attached = false;
        ownsTempDir = false;
    }

    bool XLPictureReader::isOpen() const
    {
        return !openedPath.empty();
    }

    bool XLPictureReader::ensureTempDir() const
    {
        if (!tempDir.empty()) return true;
        if (openedPath.empty()) return false;
        if (!ownsTempDir) return false;
        try {
            namespace fs = std::filesystem;
            fs::path tmp = fs::temp_directory_path() / ("minixlsx_unzip_" + std::to_string(std::chrono::system_clock::now().time_since_epoch().count()));
            fs::create_directories(tmp);
            if (cc::neolux::utils::KFZippa::unzip(openedPath, tmp.string())) {
                tempDir = tmp.string();
                return true;
            }
        } catch (...) {}
        return false;
    }

    std::string XLPictureReader::getTempDir() const
    {
        if (!ensureTempDir()) return "";
        return tempDir;
    }

    void XLPictureReader::cleanupTempDir()
    {
        if (!tempDir.empty()) {
            try {
                namespace fs = std::filesystem;
                fs::remove_all(tempDir);
            } catch (...) {}
            tempDir.clear();
        }
    }

    std::vector<PictureInfo> XLPictureReader::getPictures(unsigned int sheetIndex) const
    {
        std::vector<PictureInfo> out;
        if (!isOpen()) return out;

        // 直接构造 drawingN.xml 路径，不读取 sheet 关系
        std::string drawingPath = findDrawingPathForSheet(sheetIndex);
        if (drawingPath.empty()) return out;

        return parseDrawingXML(drawingPath);
    }

    std::vector<SheetPicture> XLPictureReader::getSheetPictures(unsigned int sheetIndex) const
    {
        std::vector<SheetPicture> out;
        if (!isOpen()) return out;

        // 直接构造 drawingN.xml 路径，不读取 sheet 关系
        std::string drawingPath = findDrawingPathForSheet(sheetIndex);
        if (drawingPath.empty()) return out;

        return parseDrawingXMLForSheetPictures(drawingPath);
    }

    std::optional<std::vector<uint8_t>> XLPictureReader::getPictureRaw(unsigned int sheetIndex, const std::string& ref) const
    {
        if (!isOpen()) return std::nullopt;
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
                    za.open(openedPath);
                    std::string data = za.getEntry(entry);
                    if (data.empty()) {
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

    std::string XLPictureReader::findDrawingPathForSheet(unsigned int sheetIndex) const
    {
        if (!ensureTempDir()) return "";

        namespace fs = std::filesystem;
        fs::path tmp(tempDir);

        std::string drawingFile = "drawings/drawing" + std::to_string(sheetIndex + 1) + ".xml";
        fs::path drawingPath = tmp / "xl" / drawingFile;
        if (!fs::exists(drawingPath)) {
            if (sheetIndex == 0) {
                fs::path alt = tmp / "xl" / "drawings" / "drawing1.xml";
                if (fs::exists(alt)) drawingPath = alt;
            }
        }

        if (!fs::exists(drawingPath)) return "";
        return drawingPath.string().substr(tmp.string().size() + 1);
    }

    std::vector<PictureInfo> XLPictureReader::parseDrawingXML(const std::string& drawingPath) const
    {
        std::vector<PictureInfo> out;
        if (!ensureTempDir()) return out;

        namespace fs = std::filesystem;
        fs::path tmp(tempDir);
        fs::path drawingFullPath = tmp / drawingPath;

        if (!fs::exists(drawingFullPath)) return out;

        // 读取 drawing 关系，建立图片映射
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

        std::ifstream drawingFile(drawingFullPath);
        std::string drawingContent((std::istreambuf_iterator<char>(drawingFile)), std::istreambuf_iterator<char>());
        drawingFile.close();

        size_t anchorPos = 0;
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
                    std::string imageTarget = it->second;
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

    std::vector<SheetPicture> XLPictureReader::parseDrawingXMLForSheetPictures(const std::string& drawingPath) const
    {
        std::vector<SheetPicture> out;
        if (!ensureTempDir()) return out;

        namespace fs = std::filesystem;
        fs::path tmp(tempDir);
        fs::path drawingFullPath = tmp / drawingPath;

        if (!fs::exists(drawingFullPath)) return out;

        // 读取 drawing 关系，建立图片映射
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

        std::ifstream drawingFile(drawingFullPath);
        std::string drawingContent((std::istreambuf_iterator<char>(drawingFile)), std::istreambuf_iterator<char>());
        drawingFile.close();

        size_t anchorPos = 0;
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
                    std::string imageTarget = it->second;
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
                    sp.colNum = fromCol + 1;
                    sp.relativePath = relPath + "/" + imageFileName;
                    out.push_back(sp);
                }
            }

            anchorPos = anchorEnd + 20;
        }

        return out;
    }
} // namespace cc::neolux::utils::MiniXLSX
