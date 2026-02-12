// OpenXLSXWrapper — thin adapter to 3rdparty OpenXLSX
#pragma once

#include <string>
#include <vector>
#include <optional>
#include <cstdint>
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

        std::optional<std::string> getCellValue(unsigned int sheetIndex, const std::string& ref) const;
        bool setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value);
        bool setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style);
        bool save();
        std::vector<PictureInfo> getPictures(unsigned int sheetIndex) const;
        std::vector<SheetPicture> fetchAllPicturesInSheet(const std::string& sheetName) const;
        std::optional<std::vector<uint8_t>> getPictureRaw(unsigned int sheetIndex, const std::string& ref) const;
        void cleanupTempDir();

    private:
        bool ensureTempDir() const;
        struct Impl;
        Impl* impl_;
    };

} // namespace cc::neolux::utils::MiniXLSX
