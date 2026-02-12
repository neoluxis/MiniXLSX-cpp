#include "cc/neolux/utils/MiniXLSX/MiniXLSX.hpp"
#include "cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp"
#include "cc/neolux/utils/MiniXLSX/XLPictureReader.hpp"
#include <memory>

namespace cc::neolux::utils::MiniXLSX
{
    struct MiniXLSX::Impl {
        std::unique_ptr<OpenXLSXWrapper> wrapper;
        std::unique_ptr<XLPictureReader> pictures;
    };

    MiniXLSX::MiniXLSX() : impl_(new Impl())
    {
        impl_->wrapper = std::make_unique<OpenXLSXWrapper>();
        impl_->pictures = std::make_unique<XLPictureReader>();
    }

    MiniXLSX::~MiniXLSX() { close(); delete impl_; }

    bool MiniXLSX::open(const std::string& path)
    {
        bool ok = impl_->wrapper->open(path);
        if (ok && impl_->pictures) {
            impl_->pictures->open(path);
        }
        return ok;
    }

    void MiniXLSX::close()
    {
        if (impl_->wrapper) impl_->wrapper->close();
        if (impl_->pictures) impl_->pictures->close();
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
        if (!impl_->pictures) return {};
        return impl_->pictures->getPictures(sheetIndex);
    }

} // namespace cc::neolux::utils::MiniXLSX
