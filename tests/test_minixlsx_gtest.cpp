#include <gtest/gtest.h>
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/MiniXLSX/XLCellPicture.hpp"

using namespace cc::neolux::utils::MiniXLSX;

TEST(MiniXLSX_Open, OpensValidFile) {
    XLDocument doc;
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    bool opened = false;
    for (auto p : candidates) {
        if (doc.open(p)) { opened = true; break; }
    }
    // Try absolute locations based on current working directory
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
