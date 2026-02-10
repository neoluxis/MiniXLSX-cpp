#include "cc/neolux/utils/MiniXLSX/XLDocument.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
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

    bool XLDocument::open(const std::string &xlsxPath)
    {
        if (isOpen)
        {
            close();
        }

        // Create a unique temporary directory
        auto now = std::chrono::system_clock::now();
        auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
        std::string tempDirName = "MiniXLSX_" + std::to_string(timestamp);
        tempDir = std::filesystem::temp_directory_path() / tempDirName;

        // Create the directory
        if (!std::filesystem::create_directory(tempDir))
        {
            std::cerr << "Failed to create temporary directory: " << tempDir << std::endl;
            return false;
        }

        // Unzip the XLSX file
        if (!KFZippa::unzip(xlsxPath, tempDir.string()))
        {
            std::cerr << "Failed to unzip XLSX file: " << xlsxPath << std::endl;
            std::filesystem::remove_all(tempDir); // Clean up on failure
            return false;
        }

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

        // Create a unique temporary directory
        auto now = std::chrono::system_clock::now();
        auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
        std::string tempDirName = "MiniXLSX_" + std::to_string(timestamp);
        tempDir = std::filesystem::temp_directory_path() / tempDirName;

        // Create the directory structure
        if (!std::filesystem::create_directories(tempDir))
        {
            std::cerr << "Failed to create temporary directory: " << tempDir << std::endl;
            return false;
        }

        // Create basic XLSX structure
        try
        {
            // Create [Content_Types].xml
            std::filesystem::create_directories(tempDir / "_rels");
            std::ofstream contentTypes(tempDir / "[Content_Types].xml");
            contentTypes << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
)";
            contentTypes.close();

            // Create _rels/.rels
            std::ofstream rels(tempDir / "_rels" / ".rels");
            rels << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
)";
            rels.close();

            // Create xl directory structure
            std::filesystem::create_directories(tempDir / "xl" / "_rels");
            std::filesystem::create_directories(tempDir / "xl" / "worksheets");

            // Create xl/workbook.xml
            std::ofstream workbookFile(tempDir / "xl" / "workbook.xml");
            workbookFile << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>
)";
            workbookFile.close();

            // Create xl/_rels/workbook.xml.rels
            std::ofstream workbookRels(tempDir / "xl" / "_rels" / "workbook.xml.rels");
            workbookRels << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
)";
            workbookRels.close();

            // Create xl/sharedStrings.xml
            std::ofstream sharedStrings(tempDir / "xl" / "sharedStrings.xml");
            sharedStrings << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>
)";
            sharedStrings.close();

            // Create xl/styles.xml
            std::ofstream styles(tempDir / "xl" / "styles.xml");
            styles << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
)";
            styles.close();

            // Create xl/worksheets/sheet1.xml
            std::ofstream sheet1(tempDir / "xl" / "worksheets" / "sheet1.xml");
            sheet1 << R"(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
    </sheetData>
</worksheet>
)";
            sheet1.close();

            // Initialize workbook
            isOpen = true;
            this->xlsxPath = xlsxPath;
            workbook = new XLWorkbook(*this);
            if (!workbook->load())
            {
                std::cerr << "Failed to load newly created workbook." << std::endl;
                close();
                return false;
            }
            isModified = true; // Mark as modified since we created it
            return true;
        }
        catch (const std::exception& e)
        {
            std::cerr << "Error creating XLSX structure: " << e.what() << std::endl;
            std::filesystem::remove_all(tempDir);
            return false;
        }
    }

    void XLDocument::close()
    {
        if (isOpen)
        {
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
        // Implementation for saving the document as a new XLSX file
        if (!isOpen)
        {
            std::cerr << "Document is not open. Cannot save." << std::endl;
            return false;
        }

        // Use KFZippa to zip the tempDir contents to the xlsxPath
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
            // No changes to save
            return true;
        }

        // Use KFZippa to zip the tempDir contents to the xlsxPath
        if (!KFZippa::zip(tempDir.string(), xlsxPath))
        {
            std::cerr << "Failed to zip contents to XLSX file: " << xlsxPath << std::endl;
            return false;
        }

        isModified = false;
        return true;
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