#include "msi.hpp"
#include <ctime>

using namespace GView::Type::MSI;
using namespace GView::Type::MSI::Panels;
using namespace AppCUI::Controls;
using namespace AppCUI::Utils;

// --- Helper for Date Formatting ---
static std::string TimeToString(std::time_t t)
{
    if (t == 0)
        return "";
    char buffer[64];
    std::strftime(buffer, sizeof(buffer), "%Y-%m-%d %H:%M:%S", std::localtime(&t));
    return std::string(buffer);
}

// ============================================================================
//                               INFORMATION PANEL
// ============================================================================

Information::Information(Reference<MSIFile> _msi) : TabPage("&Information")
{
    this->msi     = _msi;
    this->general = Factory::ListView::Create(this, "x:0,y:0,w:100%,h:100%", { "n:Field,w:20", "n:Value,w:60" }, ListViewFlags::None);

    UpdateGeneralInformation();
}

void Information::UpdateGeneralInformation()
{
    general->DeleteAllItems();

    auto& meta = msi->msiMeta;

    // Helper lambda to add items only if not empty
    auto add = [&](const std::string& field, const std::string& value) {
        if (!value.empty())
            general->AddItem({ field, value });
    };

    // 1. Summary Information (Metadata)
    general->AddItem("Summary Information").SetType(ListViewItem::Type::Category);
    add("Title", meta.title);
    add("Subject", meta.subject);
    add("Author", meta.author);
    add("Keywords", meta.keywords);
    add("Comments", meta.comments);
    add("Revision (UUID)", meta.revisionNumber);
    add("Creating App", meta.creatingApp);
    add("Last Saved By", meta.lastSavedBy);

    if (meta.createTime != 0)
        add("Created", TimeToString(meta.createTime));
    if (meta.lastSaveTime != 0)
        add("Last Saved", TimeToString(meta.lastSaveTime));

    // 2. Statistics
    general->AddItem("Statistics").SetType(ListViewItem::Type::Category);
    if (meta.pageCount > 0)
        add("Pages", std::to_string(meta.pageCount));
    if (meta.wordCount > 0)
        add("Words", std::to_string(meta.wordCount));

    // 3. File Technical Details
    general->AddItem("File Details").SetType(ListViewItem::Type::Category);

    std::string sizeStr;
    MSIFile::SizeToString(meta.totalSize, sizeStr);
    add("Total Size", sizeStr);

    add("Sector Size", std::to_string(msi->sectorSize) + " bytes");
    add("Mini Sector Size", std::to_string(msi->miniSectorSize) + " bytes");
}

void Information::OnAfterResize(int newWidth, int newHeight)
{
    if (general.IsValid()) {
        general->Resize(newWidth, newHeight);
    }
}

// ============================================================================
//                               TABLES PANEL
// ============================================================================

Tables::Tables(Reference<MSIFile> _msi) : TabPage("&Tables")
{
    this->msi  = _msi;
    this->list = Factory::ListView::Create(this, "x:0,y:0,w:100%,h:100%", { "n:Name,w:30", "n:Rows,w:10,a:r" }, ListViewFlags::None);

    // Set the handler to process clicks
    this->list->Handlers()->OnItemPressed = this;

    Update();
}

void Tables::Update()
{
    list->DeleteAllItems();
    const auto& dbTables = msi->GetTableList();

    for (const auto& tbl : dbTables) {
        LocalString<32> rowStr;
        if (tbl.rowCount == 0)
            rowStr.Set("-"); // Likely empty or pure schema
        else
            rowStr.Format("%u", tbl.rowCount);

        list->AddItem({ tbl.name, rowStr });
    }
}

void Tables::OnListViewItemPressed(Reference<ListView> lv, ListViewItem item)
{
    // Get the table name from the first column (index 0)
    std::string tableName = (std::string) item.GetText(0);

    // Create and show the TableViewer Dialog
    // Note: TableViewer is defined in msi.hpp (Dialogs namespace) and implemented in Dialogs.cpp
    auto viewer = new Dialogs::TableViewer(msi, tableName);
    viewer->Show();
}

void Tables::OnAfterResize(int newWidth, int newHeight)
{
    if (list.IsValid())
        list->Resize(newWidth, newHeight);
}

// ============================================================================
//                               FILES PANEL
// ============================================================================

Files::Files(Reference<MSIFile> _msi) : TabPage("&Files")
{
    this->msi  = _msi;
    this->list = Factory::ListView::Create(
          this, "x:0,y:0,w:100%,h:100%", { "n:Name,w:30", "n:Directory,w:20", "n:Component,w:20", "n:Size,w:10,a:r", "n:Version,w:15" }, ListViewFlags::None);

    Update();
}

void Files::Update()
{
    list->DeleteAllItems();
    const auto& files = msi->GetMsiFiles();

    for (const auto& f : files) {
        std::string sizeStr;
        MSIFile::SizeToString(f.Size, sizeStr);

        list->AddItem({ f.Name, f.Directory, f.Component, sizeStr, f.Version });
    }
}

void Files::OnAfterResize(int newWidth, int newHeight)
{
    if (list.IsValid())
        list->Resize(newWidth, newHeight);
}
