#include "msi.hpp"

using namespace AppCUI;
using namespace AppCUI::Utils;
using namespace AppCUI::Application;
using namespace AppCUI::Controls;
using namespace GView::Utils;
using namespace GView::Type;
using namespace GView;
using namespace GView::View;

// 1 = Transparent/Background
// w = White/Foreground
constexpr std::string_view MSI_ICON = "1111111111111111"
                                      "1111111111111111"
                                      "wwwwwwwwwwwwwwww"
                                      "1111111111111111"
                                      "1111111111111111"
                                      "w1111w1wwwww1www"
                                      "ww11ww1w111111w1"
                                      "w1ww1w1www1111w1"
                                      "w1111w111www11w1"
                                      "w1111w11111w11w1"
                                      "w1111w1wwwww1www"
                                      "1111111111111111"
                                      "1111111111111111"
                                      "wwwwwwwwwwwwwwww"
                                      "1111111111111111"
                                      "1111111111111111";

extern "C" {
PLUGIN_EXPORT bool Validate(const AppCUI::Utils::BufferView& buf, const std::string_view& extension)
{
    // Basic Header Size Check
    if (buf.GetLength() < sizeof(MSI::OLEHeader))
        return false;

    // Signature Check
    auto h = buf.GetObject<MSI::OLEHeader>();
    if (h->signature != MSI::OLE_SIGNATURE)
        return false;

    // Strict Size Calculation
    uint32 sectorSize = 1 << h->sectorShift;
    if (sectorSize < 512 || sectorSize > 4096)
        return false;

    /*uint64 minFileSize = 512;
    minFileSize += (uint64) h->numFatSectors * sectorSize;
    minFileSize += (uint64) h->numDirSectors * sectorSize;
    minFileSize += (uint64) h->numMiniFatSectors * sectorSize;

    if (buf.GetLength() < minFileSize)
        return false;*/

    return true;
}

PLUGIN_EXPORT TypeInterface* CreateInstance()
{
    return new MSI::MSIFile();
}

void CreateBufferView(Reference<WindowInterface> win, Reference<MSI::MSIFile> msi)
{
    BufferViewer::Settings settings;
    msi->UpdateBufferViewZones(settings);
    win->CreateViewer(settings);
}

PLUGIN_EXPORT bool PopulateWindow(Reference<WindowInterface> win)
{
    auto msi = win->GetObject()->GetContentType<MSI::MSIFile>();

    if (!msi->Update()) {
        AppCUI::Dialogs::MessageBox::ShowError("Error", "Failed to parse MSI file structure.");
        return false;
    }

    // Container View
    ContainerViewer::Settings settings;
    settings.SetIcon(MSI_ICON);

    settings.AddProperty("Type", "MSI (Compound File)");
    settings.AddProperty("Title", msi->msiMeta.title);
    settings.AddProperty("Author", msi->msiMeta.author);
    settings.AddProperty("Path", msi->obj->GetPath());

    std::string sizeAsString;
    MSI::MSIFile::SizeToString(msi->obj->GetData().GetSize(), sizeAsString);

    settings.AddProperty("Size", sizeAsString);

    // Columns for MSI Files
    settings.SetColumns({ "n:&Name,a:l,w:40", "n:&Directory,a:l,w:20", "n:&Component,a:l,w:20", "n:&Size,a:r,w:10", "n:&Version,a:l,w:15" });

    settings.SetEnumerateCallback(msi.ToObjectRef<ContainerViewer::EnumerateInterface>());
    settings.SetOpenItemCallback(msi.ToObjectRef<ContainerViewer::OpenItemInterface>());
    win->CreateViewer(settings);

    // Buffer View
    CreateBufferView(win, msi);

    // Panels
    win->AddPanel(Pointer<TabPage>(new MSI::Panels::Information(msi)), true);
    win->AddPanel(Pointer<TabPage>(new MSI::Panels::Files(msi)), true);
    win->AddPanel(Pointer<TabPage>(new MSI::Panels::Tables(msi)), true);

    return true;
}

PLUGIN_EXPORT void UpdateSettings(IniSection sect)
{
    sect["Pattern"]     = "magic:D0 CF 11 E0 A1 B1 1A E1";
    sect["Priority"]    = 1;
    sect["Description"] = "Windows Installer Database (*.msi)";
}
}