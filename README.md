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
