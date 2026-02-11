#include <gtest/gtest.h>
#include "cc/neolux/utils/MiniXLSX/MiniXLSX.hpp"
#include "cc/neolux/utils/MiniXLSX/Types.hpp"
#include <OpenXLSX.hpp>

using namespace cc::neolux::utils::MiniXLSX;

TEST(MiniXLSX_StyleWrite, ApplyBasicFillAndBorder)
{
    const std::string out = "test_style_write.xlsx"; // written in build dir

    // Create base file with OpenXLSX
    {
        OpenXLSX::XLDocument doc;
        doc.create(out, true);
        auto wb = doc.workbook();
        auto ws = wb.worksheet(1);
        ws.cell("A1").value() = "styled";
        doc.save();
        doc.close();
    }

    // Apply style via MiniXLSX facade
    {
        MiniXLSX api;
        ASSERT_TRUE(api.open(out));
        CellStyle style;
        style.backgroundColor = "#FFFF00"; // yellow
        style.border = CellBorderStyle::Thin;
        style.borderColor = "#FF0000"; // red
        ASSERT_TRUE(api.setCellStyle(0, "A1", style));
        ASSERT_TRUE(api.save());
        api.close();
    }

    // Reopen with OpenXLSX and verify style
    {
        OpenXLSX::XLDocument doc;
        doc.open(out);
        auto ws = doc.workbook().worksheet(1);
        auto styleIndex = ws.cell("A1").cellFormat();
        ASSERT_NE(styleIndex, OpenXLSX::XLInvalidStyleIndex);
        auto& styles = doc.styles();
        auto cf = styles.cellFormats().cellFormatByIndex(styleIndex);
        EXPECT_TRUE(cf.applyFill());
        EXPECT_TRUE(cf.applyBorder());
        doc.close();
    }
}
