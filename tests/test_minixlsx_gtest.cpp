#include <gtest/gtest.h>
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"

using namespace cc::neolux::utils::MiniXLSX;

TEST(MiniXLSX_Open, OpensValidFile) {
    XLDocument doc;
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    bool opened = false;
    for (auto p : candidates) {
        if (doc.open(p)) { opened = true; break; }
    }
    // 尝试基于当前工作目录的绝对路径
    if (!opened) {
        try {
            auto cwd = std::filesystem::current_path();
            std::vector<std::filesystem::path> absCandidates = {cwd / "test.xlsx", cwd / "build" / "test.xlsx"};
            for (auto &ap : absCandidates) {
                if (doc.open(ap.string())) { opened = true; break; }
            }
        } catch (...) {}
    }
    if (!opened) {
        GTEST_SKIP() << "test.xlsx not available, skipping";
    }
    EXPECT_TRUE(doc.isOpened());
    doc.close();
    EXPECT_FALSE(doc.isOpened());
}

TEST(MiniXLSX_Read, ReadBasicCells) {
    XLDocument doc;
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    bool opened = false;
    for (auto p : candidates) {
        if (doc.open(p)) { opened = true; break; }
    }
    if (!opened) {
        GTEST_SKIP() << "test.xlsx not available, skipping";
    }
    auto& wb = doc.getWorkbook();
    ASSERT_GE(wb.getSheetCount(), 1u);
    auto& sheet = wb.getSheet(0);
    const XLCell* a1 = sheet.getCell("A1");
    ASSERT_NE(a1, nullptr);
    std::string val = a1->getValue();
    EXPECT_FALSE(val.empty());
    doc.close();
}

TEST(MiniXLSX_Pictures, DetectPictureG7) {
    XLDocument doc;
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    bool opened = false;
    for (auto p : candidates) {
        if (doc.open(p)) { opened = true; break; }
    }
    if (!opened) {
        GTEST_SKIP() << "test.xlsx not available, skipping";
    }
    auto& wb = doc.getWorkbook();
    ASSERT_GE(wb.getSheetCount(), 1u);
    auto& sheet = wb.getSheet(0);
    const XLCell* cell = sheet.getCell("G7");
    if (cell == nullptr) {
        GTEST_SKIP() << "G7 not present in test.xlsx, skipping";
    }
    EXPECT_EQ(cell->getType(), std::string("picture"));
    const XLCellPicture* pic = dynamic_cast<const XLCellPicture*>(cell);
    ASSERT_NE(pic, nullptr);
    EXPECT_FALSE(pic->getImageFileName().empty());
    doc.close();
}

TEST(MiniXLSX_Wrapper, SheetIndexLookup) {
    OpenXLSXWrapper wrapper;
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    bool opened = false;
    for (auto p : candidates) {
        if (wrapper.open(p)) { opened = true; break; }
    }
    if (!opened) {
        GTEST_SKIP() << "test.xlsx not available, skipping";
    }

    // 验证可以通过名称查找索引
    unsigned int sheetCount = wrapper.sheetCount();
    ASSERT_GE(sheetCount, 1u);

    // 获取首个工作表名称
    std::string firstSheetName = wrapper.sheetName(0);
    ASSERT_FALSE(firstSheetName.empty());

    // 测试 sheetIndex
    auto indexOpt = wrapper.sheetIndex(firstSheetName);
    ASSERT_TRUE(indexOpt.has_value());
    EXPECT_EQ(indexOpt.value(), 0u);

    // 测试无效工作表名称
    auto invalidIndexOpt = wrapper.sheetIndex("NonExistentSheet");
    EXPECT_FALSE(invalidIndexOpt.has_value());

    // 测试 name -> index -> name 的往返一致性
    for (unsigned int i = 0; i < sheetCount; ++i) {
        std::string name = wrapper.sheetName(i);
        auto indexOpt = wrapper.sheetIndex(name);
        ASSERT_TRUE(indexOpt.has_value()) << "Failed to find index for sheet: " << name;
        EXPECT_EQ(indexOpt.value(), i);
    }

    wrapper.close();
}
