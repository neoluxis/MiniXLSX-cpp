#include "cc/neolux/utils/MiniXLSX/MiniXLSX.hpp"
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include <memory>

namespace cc::neolux::utils::MiniXLSX
{
    struct MiniXLSX::Impl {
        std::unique_ptr<OpenXLSXWrapper> wrapper;
    };

    MiniXLSX::MiniXLSX() : impl_(new Impl()) { impl_->wrapper = std::make_unique<OpenXLSXWrapper>(); }

    MiniXLSX::~MiniXLSX() { close(); delete impl_; }

    bool MiniXLSX::open(const std::string& path)
    {
        return impl_->wrapper->open(path);
    }

    void MiniXLSX::close()
    {
        if (impl_->wrapper) impl_->wrapper->close();
    }

    bool MiniXLSX::isOpen() const
    {
        return impl_->wrapper && impl_->wrapper->isOpen();
    }

    std::optional<std::string> MiniXLSX::getCellValue(unsigned int sheetIndex, const std::string& ref) const
    {
        return impl_->wrapper->getCellValue(sheetIndex, ref);
    }

    bool MiniXLSX::setCellValue(unsigned int sheetIndex, const std::string& ref, const std::string& value)
    {
        return impl_->wrapper->setCellValue(sheetIndex, ref, value);
    }

    bool MiniXLSX::setCellStyle(unsigned int sheetIndex, const std::string& ref, const CellStyle& style)
    {
        return impl_->wrapper->setCellStyle(sheetIndex, ref, style);
    }

    bool MiniXLSX::save()
    {
        return impl_->wrapper->save();
    }

    std::vector<PictureInfo> MiniXLSX::getPictures(unsigned int sheetIndex) const
    {
        return impl_->wrapper->getPictures(sheetIndex);
    }

} // namespace cc::neolux::utils::MiniXLSX
