#include <iostream>
#include <fstream>
#include <vector>
#include <string>
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include "OpenXLSX.hpp"
#include "cc/neolux/utils/MiniXLSX/XLPictureReader.hpp"

#include <cstdint>

int main(int argc, char **argv)
{
    if (argc < 3)
    {
        std::cerr << "Usage: " << argv[0] << " <xlsx_file_path> <sheet_name>" << std::endl;
        std::cerr << "Example: " << argv[0] << " test.xlsx SO13(DNo.1)MS" << std::endl;
        return 1;
    }

    std::string path = argv[1];
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper wrapper;

    if (!wrapper.open(path))
    {
        std::cerr << "Failed to open input xlsx: " << path << std::endl;
        return 2;
    }

    // 通过名称查找工作表索引
    std::string targetSheet = argv[2];
    std::cout << "Looking for sheet: " << targetSheet << std::endl;
    int sheetIndex = -1;
    for (unsigned int i = 0; i < wrapper.sheetCount(); ++i)
    {
        if (wrapper.sheetName(i) == targetSheet)
        {
            sheetIndex = static_cast<int>(i);
            break;
        }
    }
    if (sheetIndex < 0)
    {
        std::cerr << "Sheet '" << targetSheet << "' not found" << std::endl;
        wrapper.close();
        return 3;
    }

    // 获取 G7 图片字节并输出尺寸
    // auto rawOpt = wrapper.getPictureRaw(static_cast<unsigned int>(sheetIndex), "G7");
    // if (rawOpt.has_value()) {
    //     const std::vector<uint8_t> &bytes = rawOpt.value();

    //     auto getPngSize = [](const std::vector<uint8_t>& b, int &w, int &h) -> bool {
    //         // PNG signature
    //         if (b.size() < 24) return false;
    //         const unsigned char pngSig[8] = {0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};
    //         if (!std::equal(pngSig, pngSig+8, b.begin())) return false;
    //         // IHDR starts at offset 8+4 (length) + 4 ('IHDR') => offset 16 for width
    //         w = (b[16]<<24) | (b[17]<<16) | (b[18]<<8) | b[19];
    //         h = (b[20]<<24) | (b[21]<<16) | (b[22]<<8) | b[23];
    //         return true;
    //     };

    //     auto getJpegSize = [](const std::vector<uint8_t>& b, int &w, int &h) -> bool {
    //         if (b.size() < 4) return false;
    //         size_t i = 2;
    //         while (i + 1 < b.size()) {
    //             if (b[i] != 0xFF) { ++i; continue; }
    //             uint8_t marker = b[i+1];
    //             i += 2;
    //             if (marker == 0xD9 || marker == 0xDA) break;
    //             if (i + 1 >= b.size()) break;
    //             uint16_t length = (b[i]<<8) | b[i+1];
    //             if (length < 2) return false;
    //             if (marker >= 0xC0 && marker <= 0xC3) {
    //                 if (i + 5 >= b.size()) return false;
    //                 // precision = b[i+2]
    //                 h = (b[i+3]<<8) | b[i+4];
    //                 w = (b[i+5]<<8) | b[i+6];
    //                 return true;
    //             }
    //             i += length;
    //         }
    //         return false;
    //     };

    //     int w=0,h=0;
    //     bool got = false;
    //     if (getPngSize(bytes, w, h)) got = true;
    //     else if (getJpegSize(bytes, w, h)) got = true;

    //     if (got) std::cout << "G7 image size: " << w << "x" << h << std::endl;
    //     else std::cout << "G7 image: " << bytes.size() << " bytes (unknown format)" << std::endl;
    // } else {
    //     std::cerr << "No picture found at G7 on sheet " << targetSheet << std::endl;
    // }

    // 获取工作表图片列表
    auto allPics = wrapper.fetchAllPicturesInSheet(targetSheet);
    std::cout << "Found " << allPics.size() << " pictures in sheet '" << targetSheet << "':" << std::endl;
    for (const auto &pic : allPics)
    {
        std::cout << "  Row: " << pic.row << ", Col: " << pic.col
                  << " (Num: " << pic.rowNum << "," << pic.colNum << ") "
                  << "Path: " << pic.relativePath << std::endl;
    }

    // 测试：写入 G7 并保存
    std::cout << "Setting text '[D]' in cell G7 (empty cell, not containing a picture) and SAVING..." << std::endl;
    if (!wrapper.setCellValue(static_cast<unsigned int>(sheetIndex), "G7", "[D]"))
    cc::neolux::utils::MiniXLSX::XLPictureReader pictures;
    if (!pictures.open(path)) {
        std::cerr << "Failed to prepare picture reader for: " << path << std::endl;
        wrapper.close();
        return 4;
    }

    {
        std::cerr << "Failed to set cell value for G7" << std::endl;
    }
    else
    {
        std::cout << "Successfully set cell value. Now saving the file..." << std::endl;
        if (wrapper.save())
        {
            std::cout << "File saved successfully!" << std::endl;
        }
        else
        {
            std::cerr << "Failed to save file!" << std::endl;
        }
    }

    // 验证写入结果
    auto v = wrapper.getCellValue(static_cast<unsigned int>(sheetIndex), "G7");
    if (v.has_value())
        std::cout << "G7 now contains: '" << v.value() << "'" << std::endl;
    else
        std::cout << "G7 is empty" << std::endl;

    // 设置 G8 背景色为红色
    std::cout << "Setting cell G8 background color to red (#FF0000) and saving..." << std::endl;
    // load original style if any to preserve other attributes
    cc::neolux::utils::MiniXLSX::CellStyle style;
    style.backgroundColor = "#FF0000"; // red
    if (wrapper.setCellStyle(static_cast<unsigned int>(sheetIndex), "G8", style))
    {
        std::cout << "Successfully set cell style for G8. Now saving the file..." << std::endl;
        if (wrapper.save())
        {
            std::cout << "File saved successfully!" << std::endl;
        }
        else
        {
            std::cerr << "Failed to save file!" << std::endl;
        }
    }
    else
    {
        std::cerr << "Failed to set cell style for G8" << std::endl;
    }

    // 保存后重新读取图片列表
    std::cout << "After save, checking if pictures are still detectable..." << std::endl;
    auto allPicsAfterSave = wrapper.fetchAllPicturesInSheet(targetSheet);
    std::cout << "Found " << allPicsAfterSave.size() << " pictures in sheet '" << targetSheet << "' after save:" << std::endl;
    for (const auto &pic : allPicsAfterSave)
    {
        std::cout << "  Row: " << pic.row << ", Col: " << pic.col
                  << " (Num: " << pic.rowNum << "," << pic.colNum << ") "
                  << "Path: " << pic.relativePath << std::endl;
    }

    // Manually cleanup temp directory when done
    wrapper.cleanupTempDir();

    wrapper.close();
    return 0;
}
