#include "msi.hpp"
#include <algorithm>

using namespace GView::Type::MSI;
using namespace AppCUI::Utils;

// --- Static Helper Functions for Binary Parsing ---

static bool read_u32_le(const uint8_t* data, size_t avail, uint32_t& out)
{
    if (avail < 4)
        return false;
    out = (uint32_t) data[0] | ((uint32_t) data[1] << 8) | ((uint32_t) data[2] << 16) | ((uint32_t) data[3] << 24);
    return true;
}

static bool read_u64_le(const uint8_t* data, size_t avail, uint64_t& out)
{
    if (avail < 8)
        return false;
    uint32_t low, high;
    if (!read_u32_le(data, 4, low) || !read_u32_le(data + 4, avail - 4, high))
        return false;
    out = ((uint64_t) high << 32) | low;
    return true;
}

static void filetime_to_time_t(uint64_t ft, std::time_t& t)
{
    // Windows FileTime: 100ns intervals since Jan 1, 1601
    // Unix Epoch: Jan 1, 1970
    // Difference: 11644473600 seconds
    const uint64_t DIFF_SEC = 11644473600ULL;
    uint64_t seconds        = ft / 10000000ULL;
    if (seconds > DIFF_SEC)
        t = (std::time_t) (seconds - DIFF_SEC);
    else
        t = 0;
}

// --- MSIFile Implementation ---

MSIFile::MSIFile() : header{}, sectorSize{ 0 }, miniSectorSize{ 0 }
{
}

MSIFile::~MSIFile()
{
    for (auto* entry : linearDirList)
        delete entry;
    linearDirList.clear();
}

bool MSIFile::Update()
{
    OLEHeader h;
    CHECK(this->obj->GetData().Copy<OLEHeader>(0, h), false, "Failed to read OLE Header");
    CHECK(h.signature == OLE_SIGNATURE, false, "Invalid OLE Signature");

    this->header            = h;
    this->sectorSize        = 1 << header.sectorShift;
    this->miniSectorSize    = 1 << header.miniSectorShift;
    this->msiMeta.totalSize = this->obj->GetData().GetSize();

    CHECK(LoadFAT(), false, "Failed to load FAT");
    CHECK(LoadDirectory(), false, "Failed to load Directory");
    CHECK(LoadMiniFAT(), false, "Failed to load MiniFAT");

    BuildTree(this->rootDir);
    ParseSummaryInformation();

    // Database Loading (Implementation in MSIDatabase.cpp)
    if (LoadStringPool()) {
        LoadDatabase();
        LoadTables();
    }

    return true;
}

// --- OLE Core Parsing ---

bool MSIFile::LoadFAT()
{
    uint32 numSectors = header.numFatSectors;
    FAT.clear();
    // Pre-allocate to avoid reallocations
    FAT.reserve(std::max(numSectors * (sectorSize / 4), (uint32) 1024));

    std::vector<uint32> difatList;
    difatList.reserve(numSectors);

    // 1. Header DIFAT
    for (int i = 0; i < 109; i++) {
        if (header.difat[i] == ENDOFCHAIN || header.difat[i] == NOSTREAM)
            break;
        difatList.push_back(header.difat[i]);
    }

    // 2. External DIFAT
    uint32 currentDifatSector = header.firstDifatSector;
    uint32 entriesPerDifat    = (sectorSize / 4) - 1;
    uint32 safetyLimit        = 0;

    while (currentDifatSector != ENDOFCHAIN && currentDifatSector != NOSTREAM && safetyLimit++ < 10000) {
        uint64 offset = (uint64) (currentDifatSector + 1) * sectorSize;
        auto view     = this->obj->GetData().Get(offset, sectorSize, true);
        if (!view.IsValid())
            break;

        const uint32* data = reinterpret_cast<const uint32*>(view.GetData());
        for (uint32 k = 0; k < entriesPerDifat; k++) {
            if (data[k] != ENDOFCHAIN && data[k] != NOSTREAM)
                difatList.push_back(data[k]);
        }
        currentDifatSector = data[entriesPerDifat];
    }

    // 3. Load FAT Sectors
    for (uint32 sect : difatList) {
        uint64 offset = (uint64) (sect + 1) * sectorSize;
        auto view     = this->obj->GetData().Get(offset, sectorSize, true);
        if (view.IsValid()) {
            const uint32* data = reinterpret_cast<const uint32*>(view.GetData());
            uint32 count       = sectorSize / 4;
            for (uint32 k = 0; k < count; k++)
                FAT.push_back(data[k]);
        }
    }
    return true;
}

bool MSIFile::LoadDirectory()
{
    Buffer dirStream = GetStream(header.firstDirSector, 0, false);
    CHECK(dirStream.GetLength() > 0, false, "Failed to read Directory stream");

    // Directory entry size is fixed at 128 bytes
    uint32 count = (uint32) dirStream.GetLength() / 128;
    linearDirList.clear();
    linearDirList.reserve(count);

    if (count > 0) {
        const DirectoryEntryData* d = (const DirectoryEntryData*) dirStream.GetData();

        for (uint32 i = 0; i < count; i++) {
            DirEntry* e = new DirEntry();
            e->id       = i;
            e->data     = d[i];

            if (e->data.nameLength > 0) {
                size_t charCount = e->data.nameLength / 2;
                if (charCount > 32)
                    charCount = 32; // Safety clamp
                if (charCount > 0)
                    charCount--; // Strip null terminator

                e->name.assign(d[i].name, charCount);
                e->decodedName = MsiDecompressName(e->name);
            }
            linearDirList.push_back(e);
        }

        if (!linearDirList.empty())
            rootDir = *linearDirList[0];
    }
    return true;
}

bool MSIFile::LoadMiniFAT()
{
    Buffer fatData = GetStream(header.firstMiniFatSector, 0, false);
    if (fatData.GetLength() > 0) {
        const uint32* ptr = (const uint32*) fatData.GetData();
        size_t count      = fatData.GetLength() / 4;
        miniFAT.reserve(count);
        for (size_t i = 0; i < count; i++)
            miniFAT.push_back(ptr[i]);
    }

    if (rootDir.data.streamSize > 0) {
        miniStream = GetStream(rootDir.data.startingSectorLocation, rootDir.data.streamSize, false);
    }
    return true;
}

AppCUI::Utils::Buffer MSIFile::GetStream(uint32 startSector, uint64 size, bool isMini)
{
    std::vector<uint32>& table = isMini ? miniFAT : FAT;
    uint32 sSize               = isMini ? miniSectorSize : sectorSize;
    Buffer result;
    uint32 sect = startSector;

    // Safety checks
    uint32 limit = 0;
    // Allow a bit of overhead for fragmentation, but prevent infinite loops
    uint32 maxSectors = (size > 0) ? (uint32) (size / sSize) + 100 : 20000;

    while (sect != ENDOFCHAIN && sect != NOSTREAM) {
        if (sect >= table.size() || limit++ > maxSectors)
            break;

        if (isMini) {
            uint64 fileOffset = (uint64) sect * sSize;
            if (fileOffset + sSize <= miniStream.GetLength())
                result.Add(BufferView(miniStream.GetData() + fileOffset, sSize));
        } else {
            // Logical Sector 0 = Physical Sector 1 (after 512b Header)
            uint64 fileOffset = (uint64) (sect + 1) * sectorSize;
            auto chunk        = this->obj->GetData().CopyToBuffer(fileOffset, sSize);
            if (chunk.IsValid())
                result.Add(chunk);
        }

        sect = table[sect];

        // Optimization: Stop if we gathered enough data
        if (size > 0 && result.GetLength() >= size) {
            result.Resize(size);
            break;
        }
    }
    return result;
}

void MSIFile::BuildTree(DirEntry& parent)
{
    if (parent.data.childId == NOSTREAM)
        return;

    std::vector<uint32> siblingIDs;
    std::function<void(uint32)> traverse = [&](uint32 nodeId) {
        if (nodeId == NOSTREAM || nodeId >= linearDirList.size())
            return;
        DirEntry* node = linearDirList[nodeId];
        traverse(node->data.leftSiblingId);
        siblingIDs.push_back(nodeId);
        traverse(node->data.rightSiblingId);
    };

    traverse(parent.data.childId);

    parent.children.reserve(siblingIDs.size());
    for (uint32 id : siblingIDs) {
        DirEntry* src      = linearDirList[id];
        DirEntry childNode = *src;
        childNode.children.clear(); // Children will be built recursively

        if (childNode.data.objectType == 1 || childNode.data.objectType == 5) { // Storage or Root
            BuildTree(childNode);
        }
        parent.children.push_back(childNode);
    }
}

// --- Metadata & Utils ---

void MSIFile::ParseSummaryInformation()
{
    // Look for "\005SummaryInformation"
    for (auto* entry : linearDirList) {
        if (entry->name.find(u"SummaryInformation") == std::u16string::npos)
            continue;

        bool isMini = entry->data.streamSize < header.miniStreamCutoffSize;
        Buffer buf  = GetStream(entry->data.startingSectorLocation, entry->data.streamSize, isMini);
        if (buf.GetLength() < 48)
            return;

        msiMeta.totalSize   = static_cast<uint64_t>(buf.GetLength());
        const uint8_t* data = buf.GetData();
        size_t bufLen       = buf.GetLength();

        uint32_t sectionOffset = 0;
        // Read offset to first section (Property Set)
        if (!read_u32_le(data + 44, (bufLen >= 44 ? bufLen - 44 : 0), sectionOffset))
            return;

        if (sectionOffset >= bufLen)
            return;

        const uint8_t* sectionStart = data + sectionOffset;
        size_t sectionAvail         = bufLen - sectionOffset;
        if (sectionAvail < 8)
            return;

        uint32_t propertyCount = 0;
        if (!read_u32_le(sectionStart + 4, sectionAvail - 4, propertyCount))
            return;

        const uint8_t* propertyList = sectionStart + 8;
        size_t propertyListAvail    = (sectionAvail > 8) ? (sectionAvail - 8) : 0;

        // Validate count against available size
        if (propertyCount > propertyListAvail / 8)
            propertyCount = (uint32) (propertyListAvail / 8);

        for (uint32_t i = 0; i < propertyCount; ++i) {
            const uint8_t* plEntry = propertyList + (i * 8);
            uint32_t propID = 0, propOffset = 0;

            read_u32_le(plEntry, 8, propID);
            read_u32_le(plEntry + 4, 4, propOffset);

            if (propOffset >= sectionAvail)
                continue;

            const uint8_t* valuePtr = sectionStart + propOffset;
            size_t valueAvail       = sectionAvail - propOffset;
            if (valueAvail < 4)
                continue;

            uint32_t type = 0;
            read_u32_le(valuePtr, 4, type);
            type &= 0xFFFF; // Mask to VT type

            switch (type) {
            case 30: // VT_LPSTR
                if (propID == 2)
                    msiMeta.title = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 3)
                    msiMeta.subject = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 4)
                    msiMeta.author = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 5)
                    msiMeta.keywords = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 6)
                    msiMeta.comments = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 9)
                    msiMeta.revisionNumber = ParseLpstr(valuePtr, valueAvail);
                else if (propID == 18)
                    msiMeta.creatingApp = ParseLpstr(valuePtr, valueAvail);
                break;
            case 64: // VT_FILETIME
            {
                uint64_t ft = 0;
                if (read_u64_le(valuePtr + 4, valueAvail - 4, ft)) {
                    filetime_to_time_t(ft, (propID == 12 ? msiMeta.createTime : (propID == 13 ? msiMeta.lastSaveTime : msiMeta.lastPrintedTime)));
                }
                break;
            }
            case 2:                // VT_I2
                if (propID == 1) { /* Codepage logic if needed */
                }
                break;
            case 3: // VT_I4
            {
                uint32_t v = 0;
                read_u32_le(valuePtr + 4, valueAvail - 4, v);
                if (propID == 14)
                    msiMeta.pageCount = v;
                else if (propID == 15)
                    msiMeta.wordCount = v;
                else if (propID == 19)
                    msiMeta.security = v;
            } break;
            }
        }
        break; // Only parse the first section
    }
}

std::u16string MSIFile::MsiDecompressName(std::u16string_view encoded)
{
    static const char16_t charset[] = u"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz._";
    std::u16string result;
    result.reserve(encoded.length() * 2);

    for (size_t i = 0; i < encoded.length(); ++i) {
        uint16_t val = (uint16_t) encoded[i];
        if (val >= 0x3800 && val <= 0x47FF) {
            uint16_t packed = val - 0x3800;
            result += charset[packed & 0x3F];
            result += charset[(packed >> 6) & 0x3F];
        } else if (val >= 0x4800 && val <= 0x483F) {
            result += charset[val - 0x4800];
        } else if (val == 0x4840) {
            result += u'!';
        } else {
            result += (char16_t) val;
        }
    }
    return result;
}

void MSIFile::SizeToString(uint64 value, std::string& result)
{
    const char* units[] = { "Bytes", "KB", "MB", "GB", "TB" };
    int unitIndex       = 0;
    double doubleValue  = static_cast<double>(value);

    while (doubleValue >= 1024.0 && unitIndex < 4) {
        doubleValue /= 1024.0;
        unitIndex++;
    }

    char buffer[32];
    if (unitIndex == 0)
        snprintf(buffer, sizeof(buffer), "%llu %s", value, units[unitIndex]);
    else
        snprintf(buffer, sizeof(buffer), "%.2f %s", doubleValue, units[unitIndex]);

    result = buffer;
}

// --- Viewer & GView Interface ---

bool MSIFile::BeginIteration(std::u16string_view path, AppCUI::Controls::TreeViewItem parent)
{
    currentIterIndex = 0;

    if (path.empty()) {
        currentViewMode = ViewMode::Root;
        return true;
    }

    if (path == u"Files") {
        currentViewMode = ViewMode::Files;
        return true;
    }

    if (path == u"Tables") {
        currentViewMode = ViewMode::Tables;
        return true;
    }

    if (path == u"Streams") {
        currentViewMode   = ViewMode::Streams;
        currentIterFolder = &rootDir;
        return true;
    }

    if (parent.IsValid()) {
        DirEntry* entry = parent.GetData<DirEntry*>();
        if (entry) {
            currentViewMode   = ViewMode::Streams;
            currentIterFolder = entry;
            return true;
        }
    }

    return false;
}

bool MSIFile::PopulateItem(AppCUI::Controls::TreeViewItem item)
{
    switch (currentViewMode) {
    case ViewMode::Root:
        if (currentIterIndex == 0) {
            item.SetText(0, "Streams");
            item.SetText(1, "Folder");
            item.SetExpandable(true);
            item.SetData<DirEntry>(nullptr); // Virtual
            currentIterIndex++;
            return true;
        }
        if (currentIterIndex == 1) {
            item.SetText(0, "Files");
            item.SetText(1, "Folder");
            item.SetExpandable(true);
            item.SetData<DirEntry>(nullptr); // Virtual
            currentIterIndex++;
            return true;
        }
        if (currentIterIndex == 2) {
            item.SetText(0, "Tables");
            item.SetText(1, "Folder");
            item.SetExpandable(true);
            item.SetData<DirEntry>(nullptr); // Virtual
            currentIterIndex++;
            return true;
        }
        break;

    case ViewMode::Files:
        if (currentIterIndex < msiFiles.size()) {
            const auto& file = msiFiles[currentIterIndex];
            item.SetText(0, file.Name);
            item.SetText(1, file.Directory);
            item.SetText(2, file.Component);

            std::string sizeStr;
            SizeToString(file.Size, sizeStr);
            item.SetText(3, sizeStr);
            item.SetText(4, file.Version);

            item.SetData<DirEntry>(nullptr); // It's a file, not a stream entry
            item.SetExpandable(false);       // Files are leaves

            currentIterIndex++;
            return true;
        }
        break;

    case ViewMode::Tables:
        if (currentIterIndex < tables.size()) {
            const auto& tbl = tables[currentIterIndex];
            item.SetText(0, tbl.name);
            item.SetText(1, "Table");
            item.SetText(2, "");
            item.SetText(3, std::to_string(tbl.rowCount) + " rows");

            item.SetData<DirEntry>(nullptr);
            item.SetExpandable(false);

            currentIterIndex++;
            return true;
        }
        break;

    case ViewMode::Streams:
        if (currentIterFolder && currentIterIndex < currentIterFolder->children.size()) {
            DirEntry* child = &currentIterFolder->children[currentIterIndex];
            item.SetText(0, child->decodedName);

            if (child->data.objectType == 1 || child->data.objectType == 5) { // Storage/Root
                item.SetText(1, "Folder");
                item.SetExpandable(true);
            } else {
                item.SetText(1, "Stream");
                std::string s;
                SizeToString(child->data.streamSize, s);
                item.SetText(2, ""); // Component empty
                item.SetText(3, s);  // Size column
                item.SetExpandable(false);
            }

            item.SetData<DirEntry>(child);
            currentIterIndex++;
            return true;
        }
        break;
    }
    return false;
}

void MSIFile::OnOpenItem(std::u16string_view path, AppCUI::Controls::TreeViewItem item)
{
    // Handle opening a Table
    if (path.starts_with(u"Tables/") || path.starts_with(u"Tables\\")) {
        // item.GetText(0) contains the name.
        std::u16string txt = item.GetText(0);
        std::string tableName(txt.begin(), txt.end());

        // Show the dialog
        auto viewer = new Dialogs::TableViewer(this, tableName);
        viewer->Show();
        return;
    }

    // Handle opening a Stream
    auto e = item.GetData<DirEntry>();
    if (e && e->data.objectType == 2) {
        bool isMini    = e->data.streamSize < header.miniStreamCutoffSize;
        Buffer content = GetStream(e->data.startingSectorLocation, e->data.streamSize, isMini);
        GView::App::OpenBuffer(content, e->decodedName, "", GView::App::OpenMethod::BestMatch, "bin");
    }
}

GView::Utils::JsonBuilderInterface* MSIFile::GetSmartAssistantContext(const std::string_view& prompt, std::string_view displayPrompt)
{
    auto builder = GView::Utils::JsonBuilderInterface::Create();
    builder->AddString("Type", "Microsoft Installer (MSI)");
    builder->AddUInt("StreamsCount", (uint32) linearDirList.size());
    builder->AddUInt("TablesCount", (uint32) tables.size());
    builder->AddString("ProductTitle", msiMeta.title);
    return builder;
}

void MSIFile::UpdateBufferViewZones(GView::View::BufferViewer::Settings& settings)
{
    struct SectorTranslator : public GView::View::BufferViewer::OffsetTranslateInterface {
        uint32 sSize;
        SectorTranslator(uint32 size) : sSize(size)
        {
        }
        uint64_t TranslateToFileOffset(uint64 value, uint32) override
        {
            return (uint64_t) (value + 1) * sSize;
        }
        uint64_t TranslateFromFileOffset(uint64 value, uint32) override
        {
            return (value < 512) ? 0 : (value / sSize) - 1;
        }
    };

    settings.SetName("MSI Structure");
    settings.SetEndianess(GView::Dissasembly::Endianess::Little);
    settings.SetOffsetTranslationList({ "Sector" }, new SectorTranslator(this->sectorSize));

    settings.AddBookmark(0, 0); // Header
    if (header.firstDirSector != ENDOFCHAIN)
        settings.AddBookmark(1, (uint64) (header.firstDirSector + 1) * sectorSize);

    auto addMergedZones = [&](std::vector<uint32>& sectors, ColorPair col, std::string_view baseName) {
        if (sectors.empty())
            return;
        std::sort(sectors.begin(), sectors.end());

        uint64 start = sectors[0], count = 1;
        for (size_t i = 1; i < sectors.size(); i++) {
            if (sectors[i] == start + count) {
                count++;
            } else {
                settings.AddZone((start + 1) * sectorSize, count * sectorSize, col, baseName);
                start = sectors[i];
                count = 1;
            }
        }
        settings.AddZone((start + 1) * sectorSize, count * sectorSize, col, baseName);
    };

    // FAT Sectors
    std::vector<uint32> fatSectors;
    for (int i = 0; i < 109; i++)
        if (header.difat[i] < 0xFFFFFFFA)
            fatSectors.push_back(header.difat[i]);
    addMergedZones(fatSectors, { Color::Green, Color::Black }, "FAT Sector");

    // Directory
    std::vector<uint32> dirSectors;
    uint32 s = header.firstDirSector;
    while (s < FAT.size()) {
        dirSectors.push_back(s);
        s = FAT[s];
    }
    addMergedZones(dirSectors, { Color::Olive, Color::Black }, "Directory Sector");

    settings.AddZone(0, 512, { Color::White, Color::Magenta }, "Header");
}