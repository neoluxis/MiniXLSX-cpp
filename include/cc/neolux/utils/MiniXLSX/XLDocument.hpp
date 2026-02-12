#pragma once

#include <string>
#include <filesystem>
#include <map>
#include <vector>
#include "XLWorkbook.hpp"
#include "OpenXLSXWrapper.hpp"
#include "XLPictureReader.hpp"

namespace cc::neolux::utils::MiniXLSX
{

class XLDocument
{
private:
    std::filesystem::path tempDir;
    bool isOpen;
    bool isModified;
    std::string xlsxPath;
    XLWorkbook* workbook;
    std::unique_ptr<OpenXLSXWrapper> oxwrapper;
    std::unique_ptr<XLPictureReader> pictureReader;

public:
    XLDocument();
    ~XLDocument();

    // 获取 OpenXLSX 封装（可选）
    OpenXLSXWrapper* getWrapper() const;

    // 获取图片读取器（可选）
    XLPictureReader* getPictureReader() const;

    /**
        * @brief 打开 XLSX 文件并解压到临时目录。
        * @param xlsxPath XLSX 文件路径。
        * @return 成功返回 true，否则返回 false。
     */
    bool open(const std::string& xlsxPath);

    /**
        * @brief 创建一个包含基础结构的 XLSX 文件。
        * @param xlsxPath 新文件路径。
        * @return 成功返回 true，否则返回 false。
     */
    bool create(const std::string& xlsxPath);

    /**
        * @brief 关闭文档并清理临时目录。
     */
    void close();

    /**
        * @brief 安全关闭文档。
        * @return 关闭成功返回 true；若已修改且不允许关闭则返回 false。
     */
    bool close_safe();

    /**
        * @brief 判断文档是否已打开。
        * @return 已打开返回 true，否则返回 false。
     */
    bool isOpened() const;

    /**
        * @brief 获取 XLSX 解压后的临时目录路径。
        * @return 临时目录路径。
     */
    const std::filesystem::path& getTempDir() const;

    /**
        * @brief 获取当前文档的工作簿。
        * @return 工作簿引用。
     */
    XLWorkbook& getWorkbook();


    /**
        * @brief 另存为指定路径。
        * @param xlsxPath 目标 XLSX 文件路径。
        * @return 成功返回 true，否则返回 false。
     */
    bool saveAs(const std::string& xlsxPath);

    /**
        * @brief 若已修改则保存文档。
        * @return 成功返回 true，否则返回 false。
     */
    bool save();

    /**
        * @brief 标记文档已修改。
     */
    void markModified();
};

} // namespace cc::neolux::utils::MiniXLSX