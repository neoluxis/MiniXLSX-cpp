#pragma once
#include <string>

namespace cc::neolux::utils::MiniXLSX
{
    class XLDocument;

    class XLTemplate
    {
    public:
        static bool createTemplate(const std::string& path);
    };
} // namespace cc::neolux::utils::MiniXLSX
