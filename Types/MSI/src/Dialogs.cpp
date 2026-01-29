#include "msi.hpp"

using namespace GView::Type::MSI;
using namespace GView::Type::MSI::Dialogs;
using namespace AppCUI::Controls;
using namespace AppCUI::Input;

TableViewer::TableViewer(Reference<MSIFile> _msi, const std::string& tableName) : Window(tableName, "d:c,w:90%,h:80%", WindowFlags::Sizeable)
{
    // 1. Create the List View
    // We allow multiple selection just in case user wants to copy multiple rows (future feature)
    this->list = Factory::ListView::Create(this, "x:0,y:0,w:100%,h:100%", {}, ListViewFlags::AllowMultipleItemsSelection);

    // 2. Get Table Definition to setup Columns
    auto def = _msi->GetTableDefinition(tableName);
    if (def) {
        for (const auto& col : def->columns) {
            // Format string: "n:ColName,a:l,w:20"
            // n = name, a = align (left), w = width
            AppCUI::Utils::LocalString<128> colFormat;
            colFormat.Format("n:%s,a:l,w:20", col.name.c_str());

            // Adjust width slightly for Integers
            if (col.type & 0x8000) { // MSICOL_INTEGER check (simple mask)
                colFormat.Format("n:%s,a:r,w:10", col.name.c_str());
            }

            list->AddColumn(colFormat.GetText());
        }
    }

    // 3. Populate Data
    // We read all data into memory strings. For massive tables, a virtual list would be better,
    // but for typical MSI tables ( < 100k rows), this is acceptable.
    auto rows = _msi->ReadTableData(tableName);

    for (const auto& row : rows) {
        if (row.empty())
            continue;

        // Add the item using the first column's data
        auto item = list->AddItem(row[0]);

        // Set texts for subsequent columns
        // Note: ListView indices are 0-based matching our row vector
        for (size_t i = 1; i < row.size(); i++) {
            item.SetText((uint32) i, row[i]);
        }
    }

    // Focus the list so navigation works immediately
    list->SetFocus();
}

bool TableViewer::OnEvent(Reference<Control> control, Event eventType, int ID)
{
    // Handle the generic Window Close event (e.g. user presses X or Escape if configured)
    if (eventType == Event::WindowClose) {
        Exit(AppCUI::Dialogs::Result(0));
        return true;
    }

    return false;
}

bool TableViewer::OnKeyEvent(Key keyCode, char16 UnicodeChar)
{
    // Standard window key handling (allows moving focus, closing with Esc usually)
    if (Window::OnKeyEvent(keyCode, UnicodeChar))
        return true;

    // explicit Escape handling if Window doesn't catch it by default configuration
    if (keyCode == Key::Escape) {
        Exit(AppCUI::Dialogs::Result(0));
        return true;
    }

    return false;
}
