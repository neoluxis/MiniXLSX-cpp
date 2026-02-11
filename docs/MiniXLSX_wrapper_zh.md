<!-- MiniXLSX OpenXLSX wrapper 中文文档 -->
# MiniXLSX Wrapper（OpenXLSX 适配器）

本文件介绍项目中提供的轻量级 OpenXLSX 适配器 `OpenXLSXWrapper` 的用法与示例。

位置
- 头文件：[include/cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp](include/cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp)
- 实现：[src/OpenXLSXWrapper.cpp](src/OpenXLSXWrapper.cpp)

简介
-----
`OpenXLSXWrapper` 是对第三方库 OpenXLSX 的薄封装，提供了常用的读取/写入单元格、设置样式、以及提取图片的便捷接口，适合在不直接使用 OpenXLSX 复杂 API 时快速操作 Excel 文件（.xlsx）。

主要类与类型
----------------
- `OpenXLSXWrapper` — 主要类，常用方法如下：
  - `OpenXLSXWrapper()` / `~OpenXLSXWrapper()` — 构造/析构
  - `bool open(const std::string &path)` — 打开一个 XLSX 文件
  - `void close()` — 关闭当前文档
  - `bool isOpen() const` — 文档是否已打开
  - `unsigned int sheetCount() const` — 工作表数量
  - `std::string sheetName(unsigned int index) const` — 根据索引获取工作表名称（从 0 开始）
  - `std::optional<std::string> getCellValue(unsigned int sheetIndex, const std::string &ref) const` — 获取单元格字符串值（如 `"A1"`）
  - `bool setCellValue(unsigned int sheetIndex, const std::string &ref, const std::string &value)` — 设置单元格值
  - `bool setCellStyle(unsigned int sheetIndex, const std::string &ref, const CellStyle &style)` — 设置单元格样式（背景色、边框）
  - `bool save()` — 保存当前文档（若文件可写）
  - `std::vector<PictureInfo> getPictures(unsigned int sheetIndex) const` — 列出工作表中图片（返回 `PictureInfo` 列表）
  - `std::optional<std::vector<uint8_t>> getPictureRaw(unsigned int sheetIndex, const std::string &ref) const` — 获取指定单元格引用处图片的原始二进制数据

- `Types.hpp` 中相关数据结构（路径：[include/cc/neolux/utils/MiniXLSX/Types.hpp](include/cc/neolux/utils/MiniXLSX/Types.hpp)）：
  - `struct PictureInfo { std::string ref; std::string fileName; std::string relativePath; }` — 图片信息，`ref` 为单元格引用（如 `G7`）
  - `enum class CellBorderStyle { None, Thin, Medium, Thick }`
  - `struct CellStyle { std::string backgroundColor; std::string fontColor; CellBorderStyle border; std::string borderColor; }` — 颜色接受 `"#RRGGBB"`、`"RRGGBB"` 或 `"ffRRGGBB"` 等格式；`fontColor` 目前未被 wrapper 使用。

注意事项
--------
- `sheetIndex` 在 wrapper 中是以 0 为起点的索引（内部会转换为 OpenXLSX 的 1 起始索引）。
- `getCellValue` 返回 `std::optional<std::string>`，当取值失败时返回 `std::nullopt`。
- `setCellStyle` 会尝试保留已有的字体和其他格式，仅应用背景填充和边框相关属性。
- 图片提取：`getPictures` 会尝试将 XLSX 解压到临时目录并解析 drawing/relationships 来定位图片；`getPictureRaw` 可直接从压缩包中读取二进制数据作为备用。

示例
-----
1) 读取单元格示例

```cpp
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
using namespace cc::neolux::utils::MiniXLSX;

OpenXLSXWrapper w;
if (!w.open("example.xlsx")) {
    // 打开失败
}
auto v = w.getCellValue(0, "A1");
if (v) {
    std::cout << "A1 = " << *v << std::endl;
}
w.close();
```

2) 写入并设置样式示例

```cpp
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
using namespace cc::neolux::utils::MiniXLSX;

OpenXLSXWrapper w;
if (w.open("out.xlsx")) {
    w.setCellValue(0, "B2", "Hello");

    CellStyle s;
    s.backgroundColor = "#FFEEAA"; // 浅黄色
    s.border = CellBorderStyle::Thin;
    s.borderColor = "#000000";
    w.setCellStyle(0, "B2", s);

    w.save();
    w.close();
}
```

3) 列出图片并保存到文件（示例伪代码）

```cpp
#include <fstream>
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
using namespace cc::neolux::utils::MiniXLSX;

OpenXLSXWrapper w;
if (w.open("with_images.xlsx")) {
    auto pics = w.getPictures(0);
    for (const auto &p : pics) {
        auto raw = w.getPictureRaw(0, p.ref);
        if (raw) {
            std::ofstream out(p.fileName, std::ios::binary);
            out.write(reinterpret_cast<const char*>(raw->data()), raw->size());
        }
    }
    w.close();
}
```

常见问题
---------
- 无法打开文件：请确认路径与权限，并确保文件为合法的 `.xlsx` 压缩包。
- 样式没有生效：wrapper 当前仅设置填充和边框；字体相关属性由现有格式继承，某些复杂样式可能无法通过简单字段完全表达。

后续建议
---------
- 如需更细粒度的样式控制（字体大小、对齐、数字格式等），建议直接使用 OpenXLSX 原生 API 或在 wrapper 中扩展 `CellStyle`。
- 为图片提取提供更稳健的 API（如按图片索引直接读取）可作为改进项。

如果需要，我可以将此文档链接加入项目 README 或在 API_REFERENCE_zh.md 中添加交叉引用。
