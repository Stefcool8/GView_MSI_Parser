#include "msi.hpp"

using namespace GView::Type::MSI;
using namespace GView::Type::MSI::Dialogs;
using namespace AppCUI::Controls;
using namespace AppCUI::Input;

TableViewer::TableViewer(Reference<MSIFile> _msi, const std::string& tableName) : Window(tableName, "d:c,w:95%,h:80%", WindowFlags::Sizeable)
{
    // List View
    this->list = Factory::ListView::Create(this, "x:0,y:0,w:100%,h:100%", {}, ListViewFlags::AllowMultipleItemsSelection);

    // Beyond the maxColumns count, the columns will not be displayed (gui limitation)
    int colCount = 0, maxColumns = 8;

    auto def = _msi->GetTableDefinition(tableName);
    if (def) {
        for (const auto& col : def->columns) {
            if (++colCount > maxColumns)
                break;
            // Format string: "n:ColName,a:l,w:20"
            // n = name, a = align (left), w = width
            AppCUI::Utils::LocalString<128> colFormat;
            colFormat.Format("n:%s,a:l,w:20", col.name.c_str());

            if (col.type & 0x8000) { // MSICOL_INTEGER check (simple mask)
                colFormat.Format("n:%s,a:r,w:10", col.name.c_str());
            }

            list->AddColumn(colFormat.GetText());
        }
    }

    auto rows = _msi->ReadTableData(tableName);

    for (const auto& row : rows) {
        if (row.empty())
            continue;

        // Add the item using the first column's data
        auto item = list->AddItem(row[0]);

        // Set texts for subsequent columns
        for (size_t i = 1; i < row.size(); i++) {
            if (i > maxColumns - 1)
                break;
            item.SetText((uint32) i, row[i]);
        }
    }

    // Focus the list so navigation works immediately
    list->SetFocus();
}

bool TableViewer::OnEvent(Reference<Control> control, Event eventType, int ID)
{
    if (eventType == Event::WindowClose) {
        Exit(AppCUI::Dialogs::Result(0));
        return true;
    }

    return false;
}

bool TableViewer::OnKeyEvent(Key keyCode, char16 UnicodeChar)
{
    if (Window::OnKeyEvent(keyCode, UnicodeChar))
        return true;

    // explicit Escape handling if Window doesn't catch it by default configuration
    if (keyCode == Key::Escape) {
        Exit(AppCUI::Dialogs::Result(0));
        return true;
    }

    return false;
}
