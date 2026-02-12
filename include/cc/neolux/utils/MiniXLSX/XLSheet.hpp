#pragma once

#include <string>
#include <map>
#include <vector>
#include <memory>
#include "XLCell.hpp"
#include "OpenXLSXWrapper.hpp"
#include "XLPictureReader.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    class XLWorkbook;

    class XLSheet
    {
    private:
        XLWorkbook* workbook;
        std::string name;
        std::string sheetId;
        std::string rId;
        std::map<std::string, std::unique_ptr<XLCell>> cells;
        std::vector<std::string> sharedStrings;
        OpenXLSXWrapper* oxWrapper = nullptr;
        XLPictureReader* pictureReader = nullptr;
        unsigned int oxSheetIndex = 0;
        mutable bool picturesLoaded = false;

    public:
        XLSheet(XLWorkbook& wb, const std::string& n, const std::string& sid, const std::string& rid);
        XLSheet(XLWorkbook& wb, OpenXLSXWrapper* wrapper, XLPictureReader* pictureReader, unsigned int sheetIndex);
        ~XLSheet();

        /**
         * @brief 从文档加载工作表数据。
         * @return 成功返回 true，否则返回 false。
         */
        bool load();

        /**
         * @brief 获取工作表名称。
         * @return 工作表名称。
         */
        const std::string& getName() const;

        /**
         * @brief 通过单元格引用获取单元格（例如 "A1"）。
         * @param ref 单元格引用。
         * @return 单元格指针，若不存在则返回 nullptr。
         */
        const XLCell* getCell(const std::string& ref) const;

        /**
         * @brief 通过单元格引用获取值（例如 "A1"）。
         * @param ref 单元格引用。
         * @return 单元格值，若不存在则返回空字符串。
         */
        std::string getCellValue(const std::string& ref) const;

        /**
         * @brief 通过单元格引用设置值（例如 "A1"）。
         * @param ref 单元格引用。
         * @param value 要设置的值。
         * @param type 单元格类型（"str" 表示字符串，"n" 表示数值等）。
         */
        void setCellValue(const std::string& ref, const std::string& value, const std::string& type = "str");

        /**
         * @brief 将工作表数据写回 XML。
         * @return 成功返回 true，否则返回 false。
         */
        bool save();

        // 迭代器
        using iterator = std::map<std::string, std::unique_ptr<XLCell>>::const_iterator;
        iterator begin() const { return cells.begin(); }
        iterator end() const { return cells.end(); }

        // 工具函数
        static std::string columnNumberToLetter(int col);
    };

} // namespace cc::neolux::utils::MiniXLSX