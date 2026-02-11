#pragma once
#include <string>
#include <vector>

namespace cc::neolux::utils::MiniXLSX
{
    struct PictureInfo {
        std::string ref;
        std::string fileName;
        std::string relativePath;
    };

    enum class CellBorderStyle {
        None = 0,
        Thin,
        Medium,
        Thick
    };

    struct CellStyle {
        // Hex color strings. Accepts formats like "#RRGGBB" or "ffRRGGBB" or "RRGGBB".
        std::string backgroundColor; // 背景填充色
        std::string fontColor;       // 字体颜色（暂未使用）
        CellBorderStyle border = CellBorderStyle::None; // 边框样式（应用于四周）
        std::string borderColor;     // 边框颜色
    };

} // namespace cc::neolux::utils::MiniXLSX
