#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include "cc/neolux/utils/MiniXLSX/XLTemplate.hpp"
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include "cc/neolux/utils/MiniXLSX/XLPictureReader.hpp"
#include <chrono>
#include <filesystem>
#include <iostream>
#include <fstream>
#include <stdexcept>

namespace cc::neolux::utils::MiniXLSX
{

    XLDocument::XLDocument() : isOpen(false), workbook(nullptr) {}

    XLDocument::~XLDocument()
    {
        close();
    }

    OpenXLSXWrapper* XLDocument::getWrapper() const
    {
        return oxwrapper ? oxwrapper.get() : nullptr;
    }

    XLPictureReader* XLDocument::getPictureReader() const
    {
        return pictureReader ? pictureReader.get() : nullptr;
    }

    bool XLDocument::open(const std::string &xlsxPath)
    {
        if (isOpen)
        {
            close();
        }

        // 创建唯一的临时目录
        auto now = std::chrono::system_clock::now();
        auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
        std::string tempDirName = "MiniXLSX_" + std::to_string(timestamp);
        tempDir = std::filesystem::temp_directory_path() / tempDirName;

        // 创建临时目录
        if (!std::filesystem::create_directory(tempDir))
        {
            std::cerr << "Failed to create temporary directory: " << tempDir << std::endl;
            return false;
        }

        // 解压 XLSX 文件
        if (!KFZippa::unzip(xlsxPath, tempDir.string()))
        {
            std::cerr << "Failed to unzip XLSX file: " << xlsxPath << std::endl;
            std::filesystem::remove_all(tempDir); // 失败时清理
            return false;
        }

        // 同时打开 OpenXLSX 封装，便于调用其接口
        try {
            oxwrapper = std::make_unique<OpenXLSXWrapper>();
            if (!oxwrapper->open(xlsxPath)) {
                // 非致命错误，封装可选
                oxwrapper.reset();
            }
        } catch (...) { oxwrapper.reset(); }

        // 绑定图片读取器到已解压的临时目录
        try {
            pictureReader = std::make_unique<XLPictureReader>();
            pictureReader->attach(xlsxPath, tempDir.string());
        } catch (...) { pictureReader.reset(); }

        isOpen = true;
        workbook = new XLWorkbook(*this);
        if (!workbook->load())
        {
            std::cerr << "Failed to load workbook." << std::endl;
            close();
            return false;
        }
        return true;
    }

    bool XLDocument::create(const std::string& xlsxPath)
    {
        if (isOpen)
        {
            close();
        }

        XLTemplate::createTemplate(xlsxPath);
        return open(xlsxPath);
    }

    void XLDocument::close()
    {
        if (isOpen)
        {
            if (oxwrapper) {
                try { oxwrapper->close(); } catch(...) {}
                oxwrapper.reset();
            }
            if (pictureReader) {
                try { pictureReader->close(); } catch(...) {}
                pictureReader.reset();
            }
            delete workbook;
            workbook = nullptr;
            std::filesystem::remove_all(tempDir);
            isOpen = false;
        }
    }

    bool XLDocument::close_safe()
    {
        if(isModified) {
            return false;
        }
        close();
        return true;
    }

    bool XLDocument::isOpened() const
    {
        return isOpen;
    }

    const std::filesystem::path &XLDocument::getTempDir() const
    {
        return tempDir;
    }

    bool XLDocument::saveAs(const std::string &xlsxPath)
    {
        // 另存为新的 XLSX 文件
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        // 将修改写回到 XML
        if (!workbook->save())
        {
            std::cerr << "Failed to save workbook changes." << std::endl;
            return false;
        }

        // 使用 KFZippa 重新打包
        if (!KFZippa::zip(tempDir.string(), xlsxPath))
        {
            std::cerr << "Failed to zip contents to XLSX file: " << xlsxPath << std::endl;
            return false;
        }

        this->xlsxPath = xlsxPath;
        isModified = false;
        return true;
    }

    bool XLDocument::save()
    {
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        if (!isModified)
        {
            // 没有改动，无需保存
            return true;
        }

        // 将修改写回到 XML
        if (!workbook->save())
        {
            std::cerr << "Failed to save workbook changes." << std::endl;
            return false;
        }

        // 使用 KFZippa 重新打包
        if (!KFZippa::zip(tempDir.string(), xlsxPath))
        {
            std::cerr << "Failed to zip contents to XLSX file: " << xlsxPath << std::endl;
            return false;
        }

        isModified = false;
        return true;
    }

    void XLDocument::markModified()
    {
        isModified = true;
    }

    XLWorkbook& XLDocument::getWorkbook()
    {
        if (!workbook)
        {
            throw std::runtime_error("Workbook not loaded");
        }
        return *workbook;
    }

} // namespace cc::neolux::utils::MiniXLSX