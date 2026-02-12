# MiniXLSX for C++

一个 MS Excel 文件操作库，提供常用的 Excel 操作，支持 `.xlsx` 格式

操作流程：

1. 将 XLSX 文件解压至临时目录；以时间戳和文件哈希命名以防重复或干扰
2. 通过 XML 操作读取对应信息
3. 关闭文档，删除临时文件

## Build

Linux:

```bash
cmake -B build
cmake --build build
```

Windows:

```bash
# Assume that you have already MINGW or MSYS2 installed
pacman -S cmake make 
# Or GCC and CMAKE and Make for Windows installed
cmake -B build
cmake --build build
```

Cross compiling on Linux for Windows

```bash
sudo pacman -S mingw-w64-toolchain mingw-w64

cmake -B build -DCMAKE_TOOLCHAIN_FILE=`pwd`/cmake/mingw64.cmake
cmake --build build
```

## TODO

- [x] Open a xlsx file
- [x] Open a workbook/worksheet
- [x] Read data from a cell
- [x] Write data to a cell
- [x] Get mediae 
- [x] Parse mediae, enhance cell data getter and setter
- [x] Sheet name to index lookup utility function

## API Reference

### OpenXLSXWrapper Class

#### Sheet Management
- `unsigned int sheetCount() const` - Get the number of sheets in the workbook
- `std::string sheetName(unsigned int index) const` - Get sheet name by index
- `std::optional<unsigned int> sheetIndex(const std::string& sheetName) const` - Get sheet index by name

#### Cell Operations
- `std::optional<std::string> getCellValue(unsigned int sheetIndex, const std::string& ref) const` - Read cell value
- `bool setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value)` - Write cell value
- `bool setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style)` - Set cell style

#### Picture Operations
- `std::vector<PictureInfo> getPictures(unsigned int sheetIndex) const` - Get pictures in a sheet by index
- `std::vector<SheetPicture> fetchAllPicturesInSheet(const std::string& sheetName) const` - Get pictures in a sheet by name
- `std::optional<std::vector<uint8_t>> getPictureRaw(unsigned int sheetIndex, const std::string& ref) const` - Get raw picture data

#### File Operations
- `bool open(const std::string& path)` - Open XLSX file
- `void close()` - Close file and cleanup
- `bool isOpen() const` - Check if file is open
- `bool save()` - Save changes to file
- `void cleanupTempDir()` - Cleanup temporary files
