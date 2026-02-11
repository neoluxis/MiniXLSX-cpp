#include <iostream>
#include <fstream>
#include <vector>
#include <string>
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include "OpenXLSX.hpp"

#include <cstdint>

int main(int argc, char** argv)
{
    const char* candidates[] = {"test.xlsx", "build/test.xlsx", "tests/../test.xlsx"};
    std::string path;
    bool opened = false;
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper wrapper;

    for (auto p : candidates) {
        try {
            if (wrapper.open(p)) { path = p; opened = true; break; }
        } catch(...) {}
    }
    if (!opened) {
        std::cerr << "Failed to open input xlsx. Provide test.xlsx in repo root or build/" << std::endl;
        return 2;
    }

    // Find sheet index by name
    std::string targetSheet = "SO13(DNo.1)MS";
    int sheetIndex = -1;
    for (unsigned int i = 0; i < wrapper.sheetCount(); ++i) {
        if (wrapper.sheetName(i) == targetSheet) { sheetIndex = static_cast<int>(i); break; }
    }
    if (sheetIndex < 0) {
        std::cerr << "Sheet '" << targetSheet << "' not found" << std::endl;
        wrapper.close();
        return 3;
    }

    // Get raw picture bytes for G7 and print dimensions
    auto rawOpt = wrapper.getPictureRaw(static_cast<unsigned int>(sheetIndex), "G7");
    if (rawOpt.has_value()) {
        const std::vector<uint8_t> &bytes = rawOpt.value();

        auto getPngSize = [](const std::vector<uint8_t>& b, int &w, int &h) -> bool {
            // PNG signature
            if (b.size() < 24) return false;
            const unsigned char pngSig[8] = {0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};
            if (!std::equal(pngSig, pngSig+8, b.begin())) return false;
            // IHDR starts at offset 8+4 (length) + 4 ('IHDR') => offset 16 for width
            w = (b[16]<<24) | (b[17]<<16) | (b[18]<<8) | b[19];
            h = (b[20]<<24) | (b[21]<<16) | (b[22]<<8) | b[23];
            return true;
        };

        auto getJpegSize = [](const std::vector<uint8_t>& b, int &w, int &h) -> bool {
            if (b.size() < 4) return false;
            size_t i = 2;
            while (i + 1 < b.size()) {
                if (b[i] != 0xFF) { ++i; continue; }
                uint8_t marker = b[i+1];
                i += 2;
                if (marker == 0xD9 || marker == 0xDA) break;
                if (i + 1 >= b.size()) break;
                uint16_t length = (b[i]<<8) | b[i+1];
                if (length < 2) return false;
                if (marker >= 0xC0 && marker <= 0xC3) {
                    if (i + 5 >= b.size()) return false;
                    // precision = b[i+2]
                    h = (b[i+3]<<8) | b[i+4];
                    w = (b[i+5]<<8) | b[i+6];
                    return true;
                }
                i += length;
            }
            return false;
        };

        int w=0,h=0;
        bool got = false;
        if (getPngSize(bytes, w, h)) got = true;
        else if (getJpegSize(bytes, w, h)) got = true;

        if (got) std::cout << "G7 image size: " << w << "x" << h << std::endl;
        else std::cout << "G7 image: " << bytes.size() << " bytes (unknown format)" << std::endl;
    } else {
        std::cerr << "No picture found at G7 on sheet " << targetSheet << std::endl;
    }

    // Set G8 background to red and save
    cc::neolux::utils::MiniXLSX::CellStyle cs;
    cs.backgroundColor = "#FF0000"; // red
    if (!wrapper.setCellStyle(static_cast<unsigned int>(sheetIndex), "G8", cs)) {
        std::cerr << "Failed to set cell style for G8" << std::endl;
    }
    if (!wrapper.save()) {
        std::cerr << "Failed to save workbook" << std::endl;
    }

    // Print G8 value
    auto v = wrapper.getCellValue(static_cast<unsigned int>(sheetIndex), "G8");
    if (v.has_value()) std::cout << "G8=" << v.value() << std::endl;
    else std::cout << "G8 empty" << std::endl;

    wrapper.close();
    return 0;
}
