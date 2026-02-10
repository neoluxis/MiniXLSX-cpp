#include "cc/neolux/utils/MiniXLSX/XLDrawing.hpp"
#include "cc/neolux/utils/MiniXLSX/XLSheet.hpp"
#include <fstream>
#include <iostream>
#include <unordered_map>
#include <filesystem>

namespace cc::neolux::utils::MiniXLSX
{

XLDrawing::XLDrawing(const std::string& drawingPath, const std::string& drawingRelsPath)
    : drawingPath_(drawingPath), drawingRelsPath_(drawingRelsPath)
{
}

bool XLDrawing::load()
{
    // Parse drawing rels
    std::ifstream drawingRelsFile(drawingRelsPath_);
    std::unordered_map<std::string, std::string> imageMap;
    if (drawingRelsFile.is_open())
    {
        std::string drawingRelsContent((std::istreambuf_iterator<char>(drawingRelsFile)), std::istreambuf_iterator<char>());
        drawingRelsFile.close();

        size_t relPos = 0;
        while ((relPos = drawingRelsContent.find("<Relationship ", relPos)) != std::string::npos)
        {
            size_t relEnd = drawingRelsContent.find("/>", relPos);
            if (relEnd == std::string::npos) break;

            std::string relTag = drawingRelsContent.substr(relPos, relEnd - relPos + 2);

            size_t idStartRel = relTag.find("Id=\"");
            if (idStartRel != std::string::npos)
            {
                idStartRel += 4;
                size_t idEndRel = relTag.find("\"", idStartRel);
                if (idEndRel != std::string::npos)
                {
                    std::string id = relTag.substr(idStartRel, idEndRel - idStartRel);
                    size_t targetStart = relTag.find("Target=\"");
                    if (targetStart != std::string::npos)
                    {
                        targetStart += 8;
                        size_t targetEnd = relTag.find("\"", targetStart);
                        if (targetEnd != std::string::npos)
                        {
                            std::string target = relTag.substr(targetStart, targetEnd - targetStart);
                            imageMap[id] = target;
                        }
                    }
                }
            }

            relPos = relEnd + 2;
        }
    }

    // Parse drawing
    std::ifstream drawingFile(drawingPath_);
    if (!drawingFile.is_open())
    {
        return false;
    }

    std::string drawingContent((std::istreambuf_iterator<char>(drawingFile)), std::istreambuf_iterator<char>());
    drawingFile.close();

    // Parse twoCellAnchor
    size_t anchorPos = 0;
    while ((anchorPos = drawingContent.find("<xdr:twoCellAnchor", anchorPos)) != std::string::npos)
    {
        size_t anchorEnd = drawingContent.find("</xdr:twoCellAnchor>", anchorPos);
        if (anchorEnd == std::string::npos) break;

        std::string anchorTag = drawingContent.substr(anchorPos, anchorEnd - anchorPos + 20);

        // Extract from col and row
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

        // Extract embed
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
                std::string imageTarget = it->second; // e.g., "../media/image1.jpg"
                // Extract filename and relative path
                std::filesystem::path imagePath(imageTarget);
                std::string imageFileName = imagePath.filename().string();
                std::string relPath = imagePath.parent_path().string(); // "../media"
                // But we need relPath as "media" for getFullPath
                if (relPath == "../media")
                {
                    relPath = "media";
                }
                pictures_.push_back({ref, imageFileName, relPath});
            }
        }

        anchorPos = anchorEnd + 20;
    }

    return true;
}

const std::vector<PictureInfo>& XLDrawing::getPictures() const
{
    return pictures_;
}

} // namespace cc::neolux::utils::MiniXLSX