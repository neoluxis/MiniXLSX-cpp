#pragma once

#include <string>

namespace cc::neolux::utils::MiniXLSX
{

class XLCell
{
public:
    XLCell(const std::string& ref) : reference(ref) {}
    virtual ~XLCell() = default;
    virtual std::string getValue() const = 0;
    virtual std::string getType() const = 0;
    const std::string& getReference() const { return reference; }

protected:
    std::string reference;
};

} // namespace cc::neolux::utils::MiniXLSX