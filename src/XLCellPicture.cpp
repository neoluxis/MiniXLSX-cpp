#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include <filesystem>

namespace cc::neolux::utils::MiniXLSX
{

XLCellPicture::XLCellPicture(const std::string& ref, const std::string& fileName, const std::string& relPath)
    : XLCell(ref), imageFileName(fileName), relativePath(relPath)
{
}

std::string XLCellPicture::getValue() const
{
    return imageFileName;
}

std::string XLCellPicture::getType() const
{
    return "picture";
}

const std::string& XLCellPicture::getImageFileName() const
{
    return imageFileName;
}

const std::string& XLCellPicture::getRelativePath() const
{
    return relativePath;
}

std::string XLCellPicture::getFullPath(const std::string& tempDir) const
{
    std::filesystem::path fullPath = std::filesystem::path(tempDir) / "xl" / relativePath / imageFileName;
    return fullPath.string();
}

} // namespace cc::neolux::utils::MiniXLSX