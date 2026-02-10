#pragma once

#include <string>
#include <vector>

namespace cc::neolux::utils::MiniXLSX
{

struct PictureInfo
{
    std::string ref; // e.g., "G7"
    std::string imageFileName; // e.g., "image1.jpg"
    std::string relativePath; // e.g., "media"
};

class XLDrawing
{
public:
    XLDrawing(const std::string& drawingPath, const std::string& drawingRelsPath);

    bool load();

    const std::vector<PictureInfo>& getPictures() const;

private:
    std::string drawingPath_;
    std::string drawingRelsPath_;
    std::vector<PictureInfo> pictures_;
};

} // namespace cc::neolux::utils::MiniXLSX