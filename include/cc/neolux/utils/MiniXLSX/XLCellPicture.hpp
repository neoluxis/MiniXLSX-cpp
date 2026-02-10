#pragma once

#include "XLCell.hpp"
#include <string>

namespace cc::neolux::utils::MiniXLSX
{

class XLCellPicture : public XLCell
{
public:
    XLCellPicture(const std::string& ref, const std::string& fileName, const std::string& relPath);
    std::string getValue() const override;
    std::string getType() const override;
    const std::string& getImageFileName() const;
    const std::string& getRelativePath() const;
    std::string getFullPath(const std::string& tempDir) const;

private:
    std::string imageFileName;
    std::string relativePath;
};

} // namespace cc::neolux::utils::MiniXLSX