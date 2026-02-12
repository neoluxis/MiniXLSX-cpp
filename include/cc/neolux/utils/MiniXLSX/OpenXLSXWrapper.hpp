// OpenXLSXWrapper —— OpenXLSX 的轻量封装
#pragma once

#include <string>
#include <optional>
#include "Types.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    class OpenXLSXWrapper
    {
    public:
        OpenXLSXWrapper();
        ~OpenXLSXWrapper();

        bool open(const std::string& path);
        void close();
        bool isOpen() const;

        unsigned int sheetCount() const;
        std::string sheetName(unsigned int index) const;
        std::optional<unsigned int> sheetIndex(const std::string& sheetName) const;

        std::optional<std::string> getCellValue(unsigned int sheetIndex, const std::string& ref) const;
        bool setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value);
        bool setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style);
        bool save();

    private:
        struct Impl;
        Impl* impl_;
    };

} // namespace cc::neolux::utils::MiniXLSX
