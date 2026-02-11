# MiniXLSX API 参考（中文）

本文档为 `MiniXLSX` 库的中文 API 参考，覆盖常用的打开/读取/写入/保存以及图片处理方法，方便在项目中快速使用。

**目录**
- 简介
- 快速开始
- 主要类与接口
  - `XLDocument`
  - `XLWorkbook`
  - `XLSheet`
  - `XLCell` / `XLCellData` / `XLCellPicture`
  - `OpenXLSXWrapper`（可选）
- 图片（media）处理
- 常见示例
- 注意事项

## 简介

`MiniXLSX` 是一个基于最小依赖（利用 `KFZippa` + `pugixml`，并可包装 `OpenXLSX`）的轻量级 `.xlsx` 读写库。主要思路是将 `.xlsx` 解压到临时目录，使用 XML 解析并提供便捷接口访问单元格与嵌入资源，修改后重新打包为 `.xlsx` 文件。

库的头文件位于 `include/cc/neolux/utils/MiniXLSX/`，主要实现位于 `src/`。示例与测试在 `tests/`。

## 快速开始

基本示例：打开文件、读取单元格、写入、保存与关闭。

```cpp
#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"

using namespace cc::neolux::utils::MiniXLSX;

int main() {
    XLDocument doc;
    if (!doc.open("test.xlsx")) {
        // 处理打开失败
        return 1;
    }

    // 获取工作簿与第一个工作表
    XLWorkbook& wb = doc.getWorkbook();
    XLSheet& sheet = wb.getSheet(0);

    // 读取 A1
    std::string a1 = sheet.getCellValue("A1");

    // 写入 B2（会标记文档为已修改）
    sheet.setCellValue("B2", "123", "n");

    // 保存为新文件
    doc.saveAs("out.xlsx");

    doc.close();
    return 0;
}
```

## 主要类与接口

以下按组件列出常用方法与说明。

**`XLDocument`**（文件级别）
- 构造 / 析构：`XLDocument()` / `~XLDocument()`
- 打开：`bool open(const std::string& xlsxPath)` —— 将 `.xlsx` 解压到临时目录并加载 workbook
- 创建：`bool create(const std::string& xlsxPath)` —— 使用模板创建基本 `.xlsx` 文件并打开
- 保存：`bool save()` / `bool saveAs(const std::string& xlsxPath)` —— 保存（覆盖或另存）
- 关闭：`void close()` / `bool close_safe()` —— 关闭并清理临时目录，`close_safe` 当有未保存改动时返回 false
- 状态查询：`bool isOpened() const`
- 临时目录：`const std::filesystem::path& getTempDir() const` —— 可用于调试或直接访问媒体文件路径
- 标记修改：`void markModified()` —— 内部由 `XLSheet::setCellValue` 调用
- 获取 workbook：`XLWorkbook& getWorkbook()`

示例：
```cpp
XLDocument doc;
if (!doc.open("a.xlsx")) { /* 错误处理 */ }
auto tempDir = doc.getTempDir();
// 读取 / 修改 ...
doc.save();
doc.close();
```

**`XLWorkbook`**（工作簿）
- 构造（内部由 `XLDocument` 创建）：`XLWorkbook(XLDocument &doc)`
- 加载：`bool load()`
- 保存：`bool save()` —— 调用各 sheet 的 `save()`
- 工作表数：`size_t getSheetCount() const`
- 工作表名：`const std::string& getSheetName(size_t index) const`
- 获取 sheet：`XLSheet& getSheet(size_t index)`
- 获取所属文档：`XLDocument& getDocument() const`

示例：
```cpp
XLWorkbook& wb = doc.getWorkbook();
for (size_t i = 0; i < wb.getSheetCount(); ++i) {
    std::cout << wb.getSheetName(i) << "\n";
}
```

**`XLSheet`**（工作表）
- 构造：`XLSheet(XLWorkbook& wb, const std::string& name, const std::string& sheetId, const std::string& rId)`
- 加载：`bool load()` —— 解析 `sheetData`、sharedStrings、drawing（图片）等
- 获取单元格：`const XLCell* getCell(const std::string& ref) const` —— 如 `"A1"`，返回 `XLCell*` 或 `nullptr`
- 获取单元格值：`std::string getCellValue(const std::string& ref) const` —— 方便快捷
- 设置单元格值：`void setCellValue(const std::string& ref, const std::string& value, const std::string& type = "str")` —— type 如 `"str"`、`"n"`、或共享字符串标记
- 保存：`bool save()` —— 将内存中的单元格写回 XML
- 迭代器支持：可通过 `begin()`/`end()` 访问所有单元格
- 辅助：`static std::string columnNumberToLetter(int col)`

示例（读/写）：
```cpp
XLSheet& sheet = wb.getSheet(0);
std::string v = sheet.getCellValue("C3");
sheet.setCellValue("C4", "hello", "str");
```

**`XLCell` / `XLCellData` / `XLCellPicture`**
- `XLCell`：抽象基类，包含 `getValue()`、`getType()`、`getReference()`。
- `XLCellData`：表示数据型单元格，方法：`getValue()`、`getType()`、`setValue()`、`setType()`。
- `XLCellPicture`：表示图片单元格，方法：
  - `getImageFileName()` —— 媒体文件名，如 `image1.jpg`
  - `getRelativePath()` —— 相对路径（在解析过程可能为 `media`）
  - `getFullPath(const std::string& tempDir)` —— 根据解压的 `tempDir` 返回媒体的绝对路径
  - `getValue()` 返回图片文件名，`getType()` 返回 `"picture"`

示例：
```cpp
const XLCell* c = sheet.getCell("G7");
if (c && c->getType() == "picture") {
    auto pic = dynamic_cast<const XLCellPicture*>(c);
    std::string img = pic->getImageFileName();
    std::string full = pic->getFullPath(doc.getTempDir().string());
}
```

**`OpenXLSXWrapper`（可选）**
- 这是对 `3rdparty/OpenXLSX` 的薄包装（在 `include/cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp`）。用于在需要更强大功能时直接利用 `OpenXLSX` 的 API。
- 常用方法：`open(path)`、`close()`、`isOpen()`、`sheetCount()`、`sheetName(index)`、`getCellValue(sheetIndex, ref)` 等。

## 图片（media）处理

`MiniXLSX` 在解析 `worksheet` 时会识别 `drawing` 节点并解析 `xl/drawings` 文件以及 `_rels` 来找到图片媒体文件。

- 图片信息封装在 `XLDrawing::PictureInfo`（实现细节在 `src/XLDrawing.cpp`）内，最终会在 `XLSheet::load()` 中把图片信息以 `XLCellPicture` 的形式放入 `cells` 映射中，使用单元格引用（如 `G7`）作为键。
- 要取得图片的磁盘路径：先通过 `XLCellPicture::getFullPath(doc.getTempDir().string())` 获取图片在临时解压目录下的全路径，然后可以用该路径读取或拷贝图片。

示例：导出图片到输出目录
```cpp
const XLCell* c = sheet.getCell("G7");
if (c && c->getType() == "picture") {
    auto p = dynamic_cast<const XLCellPicture*>(c);
    std::string full = p->getFullPath(doc.getTempDir().string());
    // 现在拷贝 full 到目标位置（使用 std::filesystem::copy 等）
}
```

## 常见示例合集

- 创建新文件并写入数据：
```cpp
XLDocument doc;
doc.create("new.xlsx");
auto& wb = doc.getWorkbook();
auto& sheet = wb.getSheet(0);
sheet.setCellValue("A1", "示例", "str");
doc.saveAs("new_saved.xlsx");
doc.close();
```

- 批量读取一列数据：
```cpp
XLSheet& s = wb.getSheet(0);
for (int row = 1; row <= 100; ++row) {
    std::string ref = "A" + std::to_string(row);
    std::string v = s.getCellValue(ref);
    if (!v.empty()) { /* 处理 v */ }
}
```

## 注意事项

- 资源路径：`MiniXLSX` 将 `.xlsx` 解压到临时目录，媒体文件和 XML 均在该目录下，若需直接访问请使用 `doc.getTempDir()`。
- 保存流程：编辑后需调用 `doc.save()` 或 `doc.saveAs()` 才会把修改写回 `.xlsx` 文件。`save()` 仅在 `isModified` 为 true 时进行打包。
- 兼容性：库尽量使用标准 C++17/20 特性，依赖 `pugixml`（通过 pkg-config）和 `KFZippa`，以及可选的 `OpenXLSX`。请确保在 CMake 配置阶段满足这些依赖。
- 并发：当前设计以单线程使用为主，临时目录命名包含时间戳来避免冲突，但在多线程/多进程场景下请自行管理文件路径或加锁。

## 参考文件
- 源码入口：`src/`（主要实现）
- 头文件：`include/cc/neolux/utils/MiniXLSX/`（API 定义）
- 示例与测试：`tests/test_minixlsx_gtest.cpp`

如果你希望我把这份文档转换为 HTML（或放到项目主页 docsify/ Doxygen），我可以继续把它转换并添加到 CI 文档站点。要我接着做哪一项？

## 功能核查（针对你的需求）

以下根据你列出的需求逐条评估当前 `MiniXLSX` 的实现状态、已支持/未支持的细节，以及改进建议。

1) 操作 XLSX 不会与 `OpenXLSX` 互相冲突或覆盖
- 当前状态：部分安全。
    - `MiniXLSX` 自身使用 `KFZippa` + `pugixml` 直接解析解压后的 OOXML 文件（`src/` 实现），并且实现使用命名空间 `cc::neolux::utils::MiniXLSX`；`OpenXLSX` 在 `3rdparty/OpenXLSX` 中作为第三方库存在，命名空间为 `OpenXLSX`，两者源码互不覆盖。
    - 但在 CMake 配置中，如果将 `OpenXLSX` 链接到 `MiniXLSX`（例如通过 `target_link_libraries(MiniXLSX OpenXLSX)`），会把 OpenXLSX 的符号带入最终链接产物，可能导致体积增长或符号冲突（极少见，除非两个库导出相同的全局符号）。目前仓库里 `OpenXLSX` 被构建为静态库（`.a`），且 `libOpenXLSX.a` 保留在 `build-*` 下，`libMiniXLSX` 体积较小（表明链接器仅提取了被需要的符号）。

- 建议：
    - 如果你希望 `MiniXLSX` 独立、轻量：从 `MiniXLSX` 的 CMake 中移除对 `OpenXLSX` 的默认链接（只保留 `3rdparty/OpenXLSX` 作为可选子模块），并提供一个 CMake 选项（例如 `MINIXLSX_USE_OPENXLSX=ON`）来启用链接。这样不会在未请求时把 OpenXLSX 引入最终产物。
    - 如果需要同时使用两套 API，在同一进程内并不会自动覆盖文件内容（都操作的是解压的临时目录或各自实例），但要避免并行修改同一 XLSX 的同一副本。

2) 能否获取、改写表格中的数值数据
- 当前状态：已支持（读取与写入基本数据）。
    - 读取：`XLSheet::getCell` / `getCellValue` 能读取数值与字符串；`XLCellData` 会根据类型（`t` 属性）处理共享字符串索引（type == "s" 时把 value 当索引解析）。
    - 写入：`XLSheet::setCellValue(ref, value, type)` 会在内存 `cells` 映射中新增或更新 `XLCellData`，并调用 `workbook->getDocument().markModified()` 标记修改；`XLSheet::save()` 会将这些单元格写回对应的 worksheet XML。`XLDocument::saveAs()`/`save()` 会把临时目录打包回 xlsx。

- 限制（重要）：
    - 共享字符串（`xl/sharedStrings.xml`）未实现写入/更新：若你对单元格写入字符串，库当前不会把字符串加入 `sharedStrings.xml` 并把单元格写成 `t="s"` + 索引。当前实现会把字符串以带 `t` 属性（写入时用传入的 type）直接写入 `<c t="..."><v>...</v></c>`。这可能导致与 Excel 的最佳实践不一致，或在某些情况下产生无效/不兼容的 OOXML（例如传入 `type="str"` 并非标准 `s`/`inlineStr`）。

- 建议：实现写入共享字符串管理：
    - 在保存流程中维护/重写 `xl/sharedStrings.xml`：当写入字符串单元格时，先查找/新增字符串到 sharedStrings 表并返回索引；在单元格 XML 使用 `t="s"` 并把 `<v>` 设为索引。或支持 `inlineStr`（`<is><t>...</t></is>`）以减少对 sharedStrings 的修改量。

3) 能否获取表格中的图片等文件数据
- 当前状态：读取图片：已支持；写入/新增图片：未完备。
    - 读取：`XLDrawing` 解析 `xl/drawings/drawing*.xml` 与对应的 `_rels`，并将图片信息（单元格引用、`imageFileName`、相对路径）暴露为 `XLCellPicture` 并放入 `XLSheet::cells`，可通过 `XLCellPicture::getFullPath(doc.getTempDir().string())` 直接获得磁盘上的图片路径，随后可以读取或复制图片文件。
    - 写入/新增图片：当前没有实现。要把图像写入或替换，需要：把图片放到 `xl/media/`、在 `xl/drawings/drawing*.xml` 中创建相应 `twoCellAnchor`/`blip` 节点、在 `xl/drawings/_rels/drawing*.xml.rels` 中添加 relationship、以及在工作表的 `_rels` 中添加 drawing 关系。当前库未提供这些操作的高层 API。

- 建议：若需要写入图片，需实现：
    - 在 `XLDocument`/`XLSheet` 层添加图片添加接口（例如 `addPicture(ref, imagePath)`），并在保存时：复制媒体文件到 `xl/media/`、修改 drawing XML、添加/更新 rels 文件，然后在 `XLSheet::save()` 保持 drawing 引用正确。

4) 能否设置单元格格式（颜色、线框、斜线等）
- 当前状态：不支持（读写样式能力有限或缺失）。
    - `MiniXLSX` 目前解析并操作 `sheetData`、`sharedStrings`、以及 `drawings`（图片）。但没有对 `xl/styles.xml` 的读写封装、没有 `XLStyle` 或 `setCellFormat` 的 API 来设置 `s="..."` 属性或创建新的 cellXfs 条目。

- 建议实现路线：
    - 增加 `XLStyles` 层（或借用 `OpenXLSX` 的样式实现）来解析和修改 `xl/styles.xml`。实现应包含：读取现有 `cellXfs`、`fills`、`borders`、`fonts`，并提供 API 创建/返回 `XLStyleIndex`。
    - 在 `XLSheet::setCellValue` 或 `XLCell` 增加 `setCellFormat(styleIndex)`/`setCellStyle(...)` API，保存时把 `s="<index>"` 写入单元格属性。

## 优先级建议与实现计划（供决策参考）
- 优先级高（立即实现可满足大多数需求）
    1. 修复/完善字符串写入（sharedStrings 的写入与维护）。
    2. 将写入时使用的 cell type 统一为标准值（`s` / `n` / `inlineStr`），并在 `XLSheet::save()` 中产生合规 XML（避免使用非标准 `t="str"`）。

- 优先级中
    3. 图片写入支持（`addPicture` / `replacePicture`）：复制媒体文件、更新 drawing 与 rels。
    4. 提供一个 CMake 选项 `MINIXLSX_USE_OPENXLSX`：默认不强制链接 `OpenXLSX`，当需要更高级功能时用户可启用（避免默认体积膨胀）。

- 优先级低（以后扩展）
    5. 完整样式 API（`XLStyles`）：需要实现 `xl/styles.xml` 解析/生成，API 复杂但可参考 `OpenXLSX` 的实现做 wrapper。

## CMake 与集成建议
- 若你希望 `MiniXLSX` 保持轻量并且不把 `OpenXLSX` 自动带入：在 `CMakeLists.txt` 中删除或条件化 `target_link_libraries(MiniXLSX OpenXLSX)`；替代方案是在顶层定义选项：

```cmake
option(MINIXLSX_USE_OPENXLSX "Link OpenXLSX into MiniXLSX" OFF)
if(MINIXLSX_USE_OPENXLSX)
    add_subdirectory(3rdparty/OpenXLSX)
    target_link_libraries(MiniXLSX PRIVATE OpenXLSX)
endif()
```

这样用户只在需要时启用 `-DMINIXLSX_USE_OPENXLSX=ON`。

## 小结
- 目前 `MiniXLSX` 可以：打开 `.xlsx`、读取表格单元格、读取图片并返回媒体文件路径、对单元格值进行内存层面的修改并将部分更改写回 worksheet XML 后打包为新的 `.xlsx`。
- 目前的主要不足：共享字符串写入、图片写入（新增/替换）、样式（颜色/边框/斜线等）写入功能尚未实现或不完备。

如需我把这些未实现项逐项实现（建议先实现 sharedStrings 写入与标准单元格类型支持，然后实现样式或图片写入），我可以为每一项编写详细实现方案并开始编码。请告诉我优先级或者直接授权我开始第一个改进（建议从 sharedStrings 写入开始）。

