#pragma once

#include "XLCell.hpp"
#include <string>
#include <vector>

namespace cc::neolux::utils::MiniXLSX
{

class XLCellData : public XLCell
{
public:
    XLCellData(const std::string& ref, const std::string& val, const std::string& typ, const std::vector<std::string>& sharedStrs);
    std::string getValue() const override;
    std::string getType() const override;
    void setValue(const std::string& val);
    void setType(const std::string& typ);

private:
    std::string value;
    std::string type;
    const std::vector<std::string>& sharedStrings;
};

} // namespace cc::neolux::utils::MiniXLSX