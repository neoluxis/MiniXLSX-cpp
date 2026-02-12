#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include <memory>
#include <iostream>

// 使用 OpenXLSX 作为底层实现
#include "OpenXLSX.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    struct OpenXLSXWrapper::Impl {
        std::unique_ptr<OpenXLSX::XLDocument> doc;
    };

    OpenXLSXWrapper::OpenXLSXWrapper() : impl_(new Impl()) {}
    OpenXLSXWrapper::~OpenXLSXWrapper() { close(); delete impl_; }

    bool OpenXLSXWrapper::open(const std::string& path)
    {
        try {
            impl_->doc = std::make_unique<OpenXLSX::XLDocument>();
            impl_->doc->open(path);
            return static_cast<bool>(impl_->doc && impl_->doc->isOpen());
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::open error: " << e.what() << std::endl;
            impl_->doc.reset();
            return false;
        }
    }

    void OpenXLSXWrapper::close()
    {
        if (impl_->doc)
        {
            try { impl_->doc->close(); } catch (...) {}
            impl_->doc.reset();
        }
    }

    bool OpenXLSXWrapper::isOpen() const
    {
        return impl_->doc && impl_->doc->isOpen();
    }

    unsigned int OpenXLSXWrapper::sheetCount() const
    {
        if (!impl_->doc) return 0;
        try {
            return impl_->doc->workbook().sheetCount();
        } catch (...) { return 0; }
    }

    std::string OpenXLSXWrapper::sheetName(unsigned int index) const
    {
        try {
            if (!impl_->doc) return std::string();
            auto names = impl_->doc->workbook().sheetNames();
            if (index < names.size()) return std::string(names[index]);
        } catch (...) {}
        return std::string();
    }

    std::optional<unsigned int> OpenXLSXWrapper::sheetIndex(const std::string& sheetName) const
    {
        if (!impl_->doc) return std::nullopt;
        try {
            auto names = impl_->doc->workbook().sheetNames();
            for (unsigned int i = 0; i < names.size(); ++i) {
                if (std::string(names[i]) == sheetName) {
                    return i;
                }
            }
        } catch (...) {}
        return std::nullopt;
    }

    std::optional<std::string> OpenXLSXWrapper::getCellValue(unsigned int sheetIndex, const std::string& ref) const
    {
        if (!impl_->doc) return std::nullopt;
        try {
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            auto cell = ws.cell(ref);
            std::string val = cell.getString();
            return std::optional<std::string>(val);
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::getCellValue error: " << e.what() << std::endl;
            return std::nullopt;
        }
    }

    bool OpenXLSXWrapper::setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value)
    {
        if (!impl_->doc) return false;
        try {
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            ws.cell(ref) = value;
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::setCellValue error: " << e.what() << std::endl;
            return false;
        }
    }

    bool OpenXLSXWrapper::setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style)
    {
        if (!impl_->doc) return false;
        try {
            auto &styles = impl_->doc->styles();

            // 创建填充
            OpenXLSX::XLFills &fills = styles.fills();
            OpenXLSX::XLStyleIndex fillIdx = fills.create();
            if (!style.backgroundColor.empty()) {
                std::string col = style.backgroundColor;
                if (!col.empty() && col[0] == '#') col = col.substr(1);
                if (col.size() == 6) col = std::string("FF") + col;
                fills[fillIdx].setPatternType(OpenXLSX::XLPatternSolid);
                fills[fillIdx].setColor(OpenXLSX::XLColor(col));
            }

            // 创建边框
            OpenXLSX::XLBorders &borders = styles.borders();
            OpenXLSX::XLStyleIndex borderIdx = borders.create();
            bool hasBorder = (style.border != CellBorderStyle::None);
            if (hasBorder) {
                OpenXLSX::XLLineStyle ls = OpenXLSX::XLLineStyleThin;
                switch (style.border) {
                    case CellBorderStyle::Thin: ls = OpenXLSX::XLLineStyleThin; break;
                    case CellBorderStyle::Medium: ls = OpenXLSX::XLLineStyleMedium; break;
                    case CellBorderStyle::Thick: ls = OpenXLSX::XLLineStyleThick; break;
                    default: ls = OpenXLSX::XLLineStyleThin; break;
                }
                std::string bcol = style.borderColor.empty() ? std::string("FF000000") : style.borderColor;
                if (!bcol.empty() && bcol[0] == '#') bcol = bcol.substr(1);
                if (bcol.size() == 6) bcol = std::string("FF") + bcol;
                OpenXLSX::XLColor bcolc(bcol);
                borders[borderIdx].setLeft(ls, bcolc);
                borders[borderIdx].setRight(ls, bcolc);
                borders[borderIdx].setTop(ls, bcolc);
                borders[borderIdx].setBottom(ls, bcolc);
            }

            // 创建单元格格式，尽量保留已有字体等属性
            OpenXLSX::XLCellFormats &cellFormats = styles.cellFormats();
            auto ws = impl_->doc->workbook().worksheet(static_cast<uint16_t>(sheetIndex + 1));
            OpenXLSX::XLStyleIndex baseFmt = 0;
            try {
                baseFmt = ws.cell(ref).cellFormat();
            } catch(...) { baseFmt = 0; }

            OpenXLSX::XLStyleIndex xf;
            try {
                if (baseFmt > 0) {
                    xf = cellFormats.create(cellFormats[baseFmt]);
                } else {
                    xf = cellFormats.create();
                }
            } catch(...) {
                xf = cellFormats.create();
            }

            if (!style.backgroundColor.empty()) cellFormats[xf].setFillIndex(fillIdx);
            if (hasBorder) cellFormats[xf].setBorderIndex(borderIdx);
            cellFormats[xf].setApplyFill(!style.backgroundColor.empty());
            if (hasBorder) cellFormats[xf].setApplyBorder(true);

            try {
                cellFormats[xf].setApplyFont(true);
            } catch(...) {}

            ws.cell(ref).setCellFormat(xf);
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::setCellStyle error: " << e.what() << std::endl;
            return false;
        }
    }

    bool OpenXLSXWrapper::save()
    {
        if (!impl_->doc) return false;
        try {
            impl_->doc->save();
            return true;
        } catch (const std::exception& e) {
            std::cerr << "OpenXLSXWrapper::save error: " << e.what() << std::endl;
            return false;
        }
    }
} // namespace cc::neolux::utils::MiniXLSX
