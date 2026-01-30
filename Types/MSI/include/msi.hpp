#pragma once

#include "GView.hpp"
#include <vector>
#include <string>
#include <map>
#include <functional>
#include <ctime>

namespace GView::Type::MSI
{
// Constants
constexpr uint64 OLE_SIGNATURE = 0xE11AB1A1E011CFD0;
constexpr uint32 FREESECT      = 0xFFFFFFFF;
constexpr uint32 ENDOFCHAIN    = 0xFFFFFFFE;
constexpr uint32 NOSTREAM      = 0xFFFFFFFF;

// Data Structures 
#pragma pack(push, 1)
struct OLEHeader {
    uint64 signature;
    uint8 clsid[16];
    uint16 minorVersion;
    uint16 majorVersion;
    uint16 byteOrder;
    uint16 sectorShift; // size of sectors in power-of-two, usually 9 or 12
    uint16 miniSectorShift; // size of mini-sectors in power-of-two
    uint8 reserved[6];
    uint32 numDirSectors;
    uint32 numFatSectors;
    uint32 firstDirSector;
    uint32 transactionSignature;
    uint32 miniStreamCutoffSize; // maximum size for a mini-stream
    uint32 firstMiniFatSector;
    uint32 numMiniFatSectors;
    uint32 firstDifatSector;
    uint32 numDifatSectors;
    uint32 difat[109];
};

struct DirectoryEntryData {
    char16 name[32];
    uint16 nameLength;
    uint8 objectType; // 0=Unknown, 1=Storage, 2=Stream, 5=Root
    uint8 colorFlag;  // 0=Red, 1=Black
    uint32 leftSiblingId;
    uint32 rightSiblingId;
    uint32 childId;
    uint8 clsid[16]; // GUID
    uint32 stateBits;
    uint64 creationTime;
    uint64 modifiedTime;
    uint32 startingSectorLocation;
    uint64 streamSize;
};
#pragma pack(pop)

struct DirEntry {
    uint32 id               = 0;
    DirectoryEntryData data = {};
    std::vector<DirEntry> children;
    std::u16string name;        // Raw name from Entry
    std::u16string decodedName; // Decoded MSI name for display
};

struct MsiFileEntry {
    std::string Name;
    std::string Directory;
    std::string Component;
    uint32 Size = 0;
    std::string Version;
};

struct MsiColumnInfo {
    std::string name;
    int type = 0;
    int size = 0;
};

struct MsiTableDef {
    std::string name;
    std::vector<MsiColumnInfo> columns;
    uint32 rowSize = 0;
};

struct MSITableInfo {
    std::string name;
    uint32 rowCount;
};

// Helper function
static bool read_u32_le(const uint8_t* data, size_t avail, uint32_t& out)
{
    if (avail < 4)
        return false;
    out = (uint32_t) data[0] | ((uint32_t) data[1] << 8) | ((uint32_t) data[2] << 16) | ((uint32_t) data[3] << 24);
    return true;
}


class MSIFile : public TypeInterface,
                public View::ContainerViewer::EnumerateInterface,
                public View::ContainerViewer::OpenItemInterface
{
  public:
    struct Metadata {
        std::string title;
        std::string subject;
        std::string author;
        std::string keywords;
        std::string comments;
        std::string revisionNumber;
        std::string creatingApp;
        std::string templateStr;
        std::string lastSavedBy;
        uint16_t codepage           = 0;
        std::time_t createTime      = 0;
        std::time_t lastSaveTime    = 0;
        std::time_t lastPrintedTime = 0;
        uint32_t pageCount          = 0;
        uint32_t wordCount          = 0;
        uint32_t characterCount     = 0;
        uint32_t security           = 0;
        uint64_t totalSize          = 0; 
    } msiMeta;

    uint32 sectorSize;
    uint32 miniSectorSize;

  private:
    OLEHeader header;

    std::vector<uint32> FAT;
    std::vector<uint32> miniFAT;

    AppCUI::Utils::Buffer miniStream;

    DirEntry rootDir;
    std::vector<DirEntry*> linearDirList;

    // Database
    std::vector<std::string> stringPool;
    std::vector<MSITableInfo> tables;
    std::vector<MsiFileEntry> msiFiles;
    std::map<std::string, MsiTableDef> tableDefs;
    uint32 stringBytes = 2;

    // Iteration State (Container Viewer)
    enum class ViewMode : uint8 { Root, Streams, Files, Tables };
    ViewMode currentViewMode    = ViewMode::Root;
    DirEntry* currentIterFolder = nullptr;
    size_t currentIterIndex     = 0;

    // Parsing Methods
    bool LoadFAT();
    bool LoadMiniFAT();
    bool LoadDirectory();
    void BuildTree(DirEntry& parent);
    AppCUI::Utils::Buffer GetStream(uint32 startSector, uint64 size, bool isMini);
    void ParseSummaryInformation();

    // Database Internal Methods
    static std::u16string MsiDecompressName(std::u16string_view encoded);
    static std::string ExtractLongFileName(const std::string& rawName);
    bool LoadStringPool();
    bool LoadTables();
    bool LoadDatabase();
    std::string GetString(uint32 index);

    // Helpers
    std::string ParseLpstr(const uint8_t* ptr, size_t avail);

  public:
    MSIFile();
    virtual ~MSIFile() override;

    bool Update();
    void UpdateBufferViewZones(GView::View::BufferViewer::Settings& settings);

    const std::vector<MSITableInfo>& GetTableList() const
    {
        return tables;
    }
    const std::vector<std::string>& GetStringPool() const
    {
        return stringPool;
    }
    const std::vector<MsiFileEntry>& GetMsiFiles() const
    {
        return msiFiles;
    }

    const MsiTableDef* GetTableDefinition(const std::string& tableName) const;
    std::vector<std::vector<AppCUI::Utils::String>> ReadTableData(const std::string& tableName);

    static void SizeToString(uint64 value, std::string& result);

    // GView Related
    virtual std::string_view GetTypeName() override
    {
        return "MSI";
    }
    virtual void RunCommand(std::string_view command) override
    {
    }
    virtual bool UpdateKeys(KeyboardControlsInterface* interface) override
    {
        return true;
    }
    virtual uint32 GetSelectionZonesCount() override
    {
        return 0;
    }
    virtual TypeInterface::SelectionZone GetSelectionZone(uint32 index) override
    {
        return { 0, 0 };
    }

    // Viewer Interface
    virtual bool BeginIteration(std::u16string_view path, AppCUI::Controls::TreeViewItem parent) override;
    virtual bool PopulateItem(AppCUI::Controls::TreeViewItem item) override;
    virtual void OnOpenItem(std::u16string_view path, AppCUI::Controls::TreeViewItem item) override;
    virtual GView::Utils::JsonBuilderInterface* GetSmartAssistantContext(const std::string_view& prompt, std::string_view displayPrompt) override;
};

// UI Namespaces
namespace Panels
{
    // Metadata
    class Information : public AppCUI::Controls::TabPage
    {
        Reference<MSIFile> msi;
        Reference<AppCUI::Controls::ListView> general;
        void UpdateGeneralInformation();
        void RecomputePanelsPositions();

      public:
        Information(Reference<MSIFile> msi);
        virtual void OnAfterResize(int newWidth, int newHeight) override;
    };

    class Tables : public AppCUI::Controls::TabPage, public AppCUI::Controls::Handlers::OnListViewItemPressedInterface
    {
        Reference<MSIFile> msi;
        Reference<AppCUI::Controls::ListView> list;

      public:
        Tables(Reference<MSIFile> msi);
        void Update();
        void OnListViewItemPressed(Reference<AppCUI::Controls::ListView> lv, AppCUI::Controls::ListViewItem item) override;
        virtual void OnAfterResize(int newWidth, int newHeight) override;
    };

    class Files : public AppCUI::Controls::TabPage
    {
        Reference<MSIFile> msi;
        Reference<AppCUI::Controls::ListView> list;

      public:
        Files(Reference<MSIFile> msi);
        void Update();
        virtual void OnAfterResize(int newWidth, int newHeight) override;
    };
} // namespace Panels

namespace Dialogs
{
    class TableViewer : public AppCUI::Controls::Window
    {
        AppCUI::Utils::Reference<AppCUI::Controls::ListView> list;

      public:
        TableViewer(AppCUI::Utils::Reference<MSIFile> msi, const std::string& tableName);
        virtual bool OnEvent(AppCUI::Utils::Reference<AppCUI::Controls::Control>, AppCUI::Controls::Event eventType, int ID) override;
        virtual bool OnKeyEvent(AppCUI::Input::Key keyCode, char16 UnicodeChar) override;
    };
} // namespace Dialogs

} // namespace GView::Type::MSI