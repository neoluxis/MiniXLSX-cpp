#pragma once

#include <string>
#include <vector>
#include "XLSheet.hpp"

namespace cc::neolux::utils::MiniXLSX
{
    class XLDocument;

    class XLWorkbook
    {
    private:
        XLDocument *document;
        std::vector<std::string> sheetNames;
        std::vector<std::string> sheetIds;
        std::vector<std::string> rIds;
        std::vector<XLSheet *> sheets;

    public:
        XLWorkbook(XLDocument &doc);
        ~XLWorkbook();

        /**
         * @brief 从文档加载工作簿数据。
         * @return 成功返回 true，否则返回 false。
         */
        bool load();

        /**
         * @brief 保存工作簿中的所有工作表。
         * @return 成功返回 true，否则返回 false。
         */
        bool save();

        /**
         * @brief 获取工作簿的工作表数量。
         * @return 工作表数量。
         */
        size_t getSheetCount() const;

        /**
         * @brief 获取指定索引的工作表名称。
         * @param index 工作表索引。
         * @return 工作表名称。
         */
        const std::string &getSheetName(size_t index) const;

        /**
         * @brief 获取指定索引的工作表。
         * @param index 工作表索引。
         * @return 工作表引用。
         */
        XLSheet &getSheet(size_t index);

        /**
         * @brief 获取关联的文档对象。
         * @return 文档引用。
         */
        XLDocument &getDocument() const;
    }; // Added semicolon here
} // namespace cc::neolux::utils::MiniXLSX