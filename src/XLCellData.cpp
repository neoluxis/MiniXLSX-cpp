#include "cc/neolux/utils/MiniXLSX/XLCellData.hpp"
#include <stdexcept>

namespace cc::neolux::utils::MiniXLSX
{

XLCellData::XLCellData(const std::string& ref, const std::string& val, const std::string& typ, const std::vector<std::string>& sharedStrs)
    : XLCell(ref), value(val), type(typ), sharedStrings(sharedStrs)
{
}

std::string XLCellData::getValue() const
{
    if (type == "s")
    {
        try
        {
            size_t index = std::stoul(value);
            if (index < sharedStrings.size())
            {
                return sharedStrings[index];
            }
        }
        catch (const std::exception&)
        {
            // 索引无效
        }
        return "";
    }
    else
    {
        return value;
    }
}

std::string XLCellData::getType() const
{
    return type;
}

void XLCellData::setValue(const std::string& val)
{
    value = val;
}

void XLCellData::setType(const std::string& typ)
{
    type = typ;
}

} // namespace cc::neolux::utils::MiniXLSX