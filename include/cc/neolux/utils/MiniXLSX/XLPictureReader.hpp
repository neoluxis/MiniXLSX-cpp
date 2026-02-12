#pragma once

#include <string>
#include <vector>
#include <optional>
#include <cstdint>
#include "Types.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    class XLPictureReader
    {
    public:
        XLPictureReader();
        ~XLPictureReader();

        bool open(const std::string& xlsxPath);
        bool attach(const std::string& xlsxPath, const std::string& tempDir);
        void close();
        bool isOpen() const;

        std::vector<PictureInfo> getPictures(unsigned int sheetIndex) const;
        std::vector<SheetPicture> getSheetPictures(unsigned int sheetIndex) const;
        std::optional<std::vector<uint8_t>> getPictureRaw(unsigned int sheetIndex, const std::string& ref) const;

        std::string getTempDir() const;
        void cleanupTempDir();

    private:
        bool ensureTempDir() const;
        std::string findDrawingPathForSheet(unsigned int sheetIndex) const;
        std::vector<PictureInfo> parseDrawingXML(const std::string& drawingPath) const;
        std::vector<SheetPicture> parseDrawingXMLForSheetPictures(const std::string& drawingPath) const;

        std::string openedPath;
        mutable std::string tempDir;
        bool ownsTempDir = false;
        bool attached = false;
    };

} // namespace cc::neolux::utils::MiniXLSX
