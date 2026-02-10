#include <iostream>
#include <fstream>
#include <vector>
#include <iomanip>

int main(int argc, char* argv[]) {
    if (argc != 2) {
        std::cerr << "Usage: " << argv[0] << " <xlsx_file>" << std::endl;
        return 1;
    }

    std::ifstream file(argv[1], std::ios::binary);
    if (!file) {
        std::cerr << "Cannot open file: " << argv[1] << std::endl;
        return 1;
    }

    std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(file), {});
    file.close();

    std::cout << "constexpr int templateSize = " << buffer.size() << ";" << std::endl;
    std::cout << "constexpr unsigned char templateData[" << buffer.size() << "] = {" << std::endl;

    for (size_t i = 0; i < buffer.size(); ++i) {
        if (i % 20 == 0) std::cout << "    ";
        std::cout << "0x" << std::hex << std::setw(2) << std::setfill('0') << static_cast<int>(buffer[i]);
        if (i < buffer.size() - 1) std::cout << ", ";
        if ((i + 1) % 20 == 0) std::cout << std::endl;
    }

    std::cout << std::endl << "};" << std::endl;
    return 0;
}