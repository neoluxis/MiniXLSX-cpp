#include "cc/neolux/utils/MiniXLSX/MiniXLSX.hpp"
#include "cc/neolux/utils/MiniXLSX/Types.hpp"
#include <OpenXLSX.hpp>
#include <iostream>

int main(int argc, const char** argv)
{
    if (argc < 2) {
        std::cout << "Usage: demo_set_style <path_to_xlsx>" << std::endl;
        return 1;
    }
    std::string path = argv[1];

    // Ensure file exists
    {
        OpenXLSX::XLDocument doc;
        doc.create(path, true);
        auto ws = doc.workbook().worksheet(1);
        ws.cell("A1").value() = "styled";
        doc.save();
        doc.close();
    }

    cc::neolux::utils::MiniXLSX::MiniXLSX api;
    if (!api.open(path)) {
        std::cerr << "Failed to open: " << path << std::endl;
        return 2;
    }

    cc::neolux::utils::MiniXLSX::CellStyle s;
    s.backgroundColor = "#C0FFC0"; // light green
    s.border = cc::neolux::utils::MiniXLSX::CellBorderStyle::Medium;
    s.borderColor = "#0000FF"; // blue

    if (!api.setCellStyle(0, "A1", s)) {
        std::cerr << "Failed to set style" << std::endl;
        return 3;
    }

    if (!api.save()) {
        std::cerr << "Failed to save" << std::endl;
        return 4;
    }

    std::cout << "Wrote style to " << path << std::endl;
    return 0;
}
