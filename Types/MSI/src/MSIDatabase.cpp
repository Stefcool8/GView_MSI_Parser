#include "msi.hpp"
#include <set>
#include <map>
#include <algorithm>
#include <vector>

using namespace GView::Type::MSI;
using namespace AppCUI::Utils;

constexpr int MSICOL_INTEGER = 1 << 15;
constexpr int MSICOL_INT2    = 1 << 10;

std::string MSIFile::GetString(uint32 index)
{
    if (index >= stringPool.size())
        return "";
    return stringPool[index];
}

std::string MSIFile::ExtractLongFileName(const std::string& rawName)
{
    auto pos = rawName.find('|');
    return (pos != std::string::npos && pos + 1 < rawName.size()) ? rawName.substr(pos + 1) : rawName;
}

bool MSIFile::LoadStringPool()
{
    DirEntry* entryPool = nullptr;
    DirEntry* entryData = nullptr;

    for (auto* e : linearDirList) {
        if (e->decodedName == u"!_StringPool")
            entryPool = e;
        else if (e->decodedName == u"!_StringData")
            entryData = e;
    }

    if (!entryPool || !entryData)
        return false;

    bool isMiniPool = entryPool->data.streamSize < header.miniStreamCutoffSize;
    Buffer bufPool  = GetStream(entryPool->data.startingSectorLocation, entryPool->data.streamSize, isMiniPool);

    bool isMiniData = entryData->data.streamSize < header.miniStreamCutoffSize;
    Buffer bufData  = GetStream(entryData->data.startingSectorLocation, entryData->data.streamSize, isMiniData);

    // Must contain at least one entry or header
    if (bufPool.GetLength() < 4)
        return false;

    stringPool.clear();
    uint32 count          = bufPool.GetLength() / 4;
    const uint16* poolPtr = (const uint16*) bufPool.GetData();
    const char* dataPtr   = (const char*) bufData.GetData();
    uint32 dataSize       = (uint32) bufData.GetLength();

    // Standard detection for Long vs Short string references in the pool
    bool highValid      = true;
    uint32 calcSizeHigh = 0;
    for (uint32 i = 1; i < count; i++) {
        uint32 len = poolPtr[i * 2 + 1];
        if (calcSizeHigh + len > dataSize) {
            highValid = false;
            break;
        }
        calcSizeHigh += len;
    }
    if (highValid && calcSizeHigh != dataSize)
        highValid = false;

    bool lowValid      = true;
    uint32 calcSizeLow = 0;
    for (uint32 i = 1; i < count; i++) {
        uint32 len = poolPtr[i * 2];
        if (calcSizeLow + len > dataSize) {
            lowValid = false;
            break;
        }
        calcSizeLow += len;
    }
    if (lowValid && calcSizeLow != dataSize)
        lowValid = false;

    bool useHighWord = true;
    if (lowValid && !highValid)
        useHighWord = false;

    stringPool.reserve(count);
    uint32 currentOffset = 0;

    for (uint32 i = 0; i < count; i++) {
        if (i == 0) {
            // Index 0 is always null/empty in MSI pools
            stringPool.push_back("");
            continue;
        }
        uint16 len = useHighWord ? poolPtr[i * 2 + 1] : poolPtr[i * 2];

        if (currentOffset + len > dataSize) {
            stringPool.push_back("<Error>");
            break;
        }

        std::string temp(dataPtr + currentOffset, len);
        while (!temp.empty() && temp.back() == '\0')
            temp.pop_back();
        stringPool.push_back(temp);
        currentOffset += len;
    }
    return stringPool.size() > 0;
}

bool MSIFile::LoadDatabase()
{
    if (stringPool.empty())
        return false;

    // Detect String Index Size (2-byte vs 3-byte)
    // Heuristic: Check if !_Columns stream size is divisible by 8 (2-byte) or 10 (3-byte)
    this->stringBytes      = 2;
    DirEntry* columnsEntry = nullptr;
    for (auto* e : linearDirList) {
        if (e->decodedName == u"!_Columns") {
            columnsEntry = e;
            break;
        }
    }

    if (columnsEntry && columnsEntry->data.streamSize > 0) {
        uint64 sz = columnsEntry->data.streamSize;
        if (sz % 10 == 0 && sz % 8 != 0)
            this->stringBytes = 3;
        else if (sz % 8 == 0 && sz % 10 != 0)
            this->stringBytes = 2;
        else
            this->stringBytes = (stringPool.size() > 65536) ? 3 : 2;
    }

    // Parse Schema from !_Columns
    // Uses the Column-Oriented logic implemented in ReadTableData equivalent
    tableDefs.clear();
    if (columnsEntry) {
        bool isMini = columnsEntry->data.streamSize < header.miniStreamCutoffSize;
        Buffer buf  = GetStream(columnsEntry->data.startingSectorLocation, columnsEntry->data.streamSize, isMini);

        if (buf.GetLength() > 0) {
            // Dimensions
            uint32 colRowSize = (this->stringBytes * 2) + 4;
            uint32 numRows    = buf.GetLength() / colRowSize;
            const uint8* ptr  = buf.GetData();

            // Block Start Offsets (Column-Oriented)
            // Block 1: Table Names (String Indices)
            uint32 startTable = 0;

            // Block 2: Column Numbers (2-byte Integers)
            uint32 startNum = startTable + (numRows * stringBytes);

            // Block 3: Column Names (String Indices)
            uint32 startName = startNum + (numRows * 2);

            // Block 4: Types (2-byte Integers)
            uint32 startType = startName + (numRows * stringBytes);

            // Iterate Rows by jumping between blocks
            for (uint32 i = 0; i < numRows; i++) {
                uint32 offset = i * colRowSize;
                // Read Table Name (From Block 1)
                uint32 offsetTable = startTable + (i * stringBytes);
                // Check bounds
                if (offsetTable + stringBytes > buf.GetLength())
                    break;

                uint32 tableIdx = 0;
                if (stringBytes == 2)
                    tableIdx = *(uint16*) (ptr + offsetTable);
                else
                    tableIdx = ptr[offsetTable] | (ptr[offsetTable + 1] << 8) | (ptr[offsetTable + 2] << 16);

                // Column Number (Block 2)
                uint32 offsetNum = startNum + (i * 2);
                uint16 colNum    = *(uint16*) (ptr + offsetNum);
                if (colNum & 0x8000) {
                    colNum &= 0x7FFF; // Remove the high bit
                }

                // Column Name (Block 3)
                uint32 offsetName = startName + (i * stringBytes);
                uint32 nameIdx    = 0;
                if (stringBytes == 2)
                    nameIdx = *(uint16*) (ptr + offsetName);
                else
                    nameIdx = ptr[offsetName] | (ptr[offsetName + 1] << 8) | (ptr[offsetName + 2] << 16);

                // Type (Block 4)
                uint32 offsetType = startType + (i * 2);
                uint16 diskType   = *(uint16*) (ptr + offsetType);

                // Strip the 0x8000 bit (Nullable/Temp flag)
                diskType &= 0x7FFF;

                // Disk Type to Internal Flags
                int finalType = 0;
                bool isString = (diskType & 0x0800) != 0;

                if (!isString) {
                    // Integer
                    finalType |= MSICOL_INTEGER;

                    // Determine size from the lower nibble
                    // 2 = i2 (2 bytes), 4 = i4 (4 bytes)
                    if ((diskType & 0x0F) == 2) {
                        finalType |= MSICOL_INT2;
                    }
                }
                // default: 0 (String).

                std::string tableNameStr = GetString(tableIdx);
                std::string colNameStr   = GetString(nameIdx);

                if (tableNameStr.empty() || tableNameStr == "<Error>")
                    continue;
                if (colNum == 0 || colNum > 255)
                    continue;

                MsiTableDef& def = tableDefs[tableNameStr];
                def.name         = tableNameStr;
                MsiColumnInfo col{ colNameStr, (int) finalType, 0 };

                if (def.columns.size() < (size_t) colNum)
                    def.columns.resize(colNum);
                def.columns[colNum - 1] = col;
            }
        }
    }
    // Column Sizes & Row Width
    for (auto& [name, def] : tableDefs) {
        uint32 rowWidth = 0;
        for (auto& col : def.columns) {
            if (col.type & MSICOL_INTEGER) {
                // 2-byte vs 4-byte
                if (col.type & MSICOL_INT2) {
                    col.size = 2;
                } else {
                    col.size = 4;
                }
            } else {
                // It is a string (flag 0x0200 was set on disk)
                col.size = stringBytes;
            }

            rowWidth += col.size;
        }
        def.rowSize = rowWidth;
    }

    // Populate Files Panel
    if (tableDefs.find("File") != tableDefs.end()) {
        msiFiles.clear();

        // Load Directories
        std::map<std::string, std::pair<std::string, std::string>> dirStructure;
        auto dirRows = ReadTableData("Directory");
        for (const auto& row : dirRows) {
            if (row.size() < 3)
                continue;
            std::string key    = std::string(row[0].GetText());
            std::string parent = std::string(row[1].GetText());
            std::string defDir = ExtractLongFileName(std::string(row[2].GetText()));
            if (!key.empty())
                dirStructure[key] = { parent, defDir };
        }

        // Load Components
        std::map<std::string, std::string> compToDir;
        auto compRows = ReadTableData("Component");
        for (const auto& row : compRows) {
            if (row.size() < 3)
                continue;
            std::string key = std::string(row[0].GetText());
            std::string dir = std::string(row[2].GetText());
            if (!key.empty())
                compToDir[key] = dir;
        }

        // Path Resolution Helper
        std::map<std::string, std::string> pathCache;
        std::function<std::string(std::string)> resolvePath = [&](std::string key) -> std::string {
            if (pathCache.find(key) != pathCache.end())
                return pathCache[key];
            if (dirStructure.find(key) == dirStructure.end())
                return key;
            auto& info = dirStructure[key];
            if (info.first.empty() || info.first == key)
                return pathCache[key] = info.second;
            std::string p         = resolvePath(info.first);
            std::string f         = info.second;
            return pathCache[key] = (p.back() == '\\') ? p + f : p + "\\" + f;
        };

        // Load Files
        auto fileRows = ReadTableData("File");
        msiFiles.reserve(fileRows.size());
        for (const auto& row : fileRows) {
            if (row.size() < 5)
                continue;
            MsiFileEntry entry;
            entry.Name      = ExtractLongFileName(std::string(row[2].GetText()));
            entry.Component = std::string(row[1].GetText());
            try {
                entry.Size = std::stoul(std::string(row[3].GetText()));
            } catch (...) {
                entry.Size = 0;
            }
            entry.Version = std::string(row[4].GetText());

            if (compToDir.find(entry.Component) != compToDir.end())
                entry.Directory = resolvePath(compToDir[entry.Component]);
            else
                entry.Directory = "<Orphaned>";
            msiFiles.push_back(entry);
        }
    }
    return true;
}

bool MSIFile::LoadTables()
{
    tables.clear();
    for (const auto& [name, def] : tableDefs) {
        uint32 count              = 0;
        std::u16string targetName = u"!" + std::u16string(name.begin(), name.end());
        DirEntry* stream          = nullptr;
        for (auto* e : linearDirList) {
            if (e->decodedName == targetName) {
                stream = e;
                break;
            }
        }

        if (stream && def.rowSize > 0)
            count = (uint32) (stream->data.streamSize / def.rowSize);
        tables.push_back({ name, count });
    }
    return true;
}

std::vector<std::vector<AppCUI::Utils::String>> MSIFile::ReadTableData(const std::string& tableName)
{
    std::vector<std::vector<AppCUI::Utils::String>> results;
    if (tableDefs.find(tableName) == tableDefs.end())
        return results;

    const auto& def           = tableDefs[tableName];
    std::u16string targetName = u"!" + std::u16string(tableName.begin(), tableName.end());
    DirEntry* tableEntry      = nullptr;
    for (auto* e : linearDirList) {
        if (e->decodedName == targetName) {
            tableEntry = e;
            break;
        }
    }

    if (!tableEntry)
        return results;

    bool isMini = tableEntry->data.streamSize < header.miniStreamCutoffSize;
    Buffer buf  = GetStream(tableEntry->data.startingSectorLocation, tableEntry->data.streamSize, isMini);
    if (def.rowSize == 0)
        return results;

    // COLUMN-ORIENTED READ LOGIC 
    uint32 numRows   = buf.GetLength() / def.rowSize;
    const uint8* ptr = buf.GetData();

    // Pre-calculate Start Offsets for each Column Block
    std::vector<uint32> colStartOffsets;
    uint32 currentOffset = 0;
    for (const auto& col : def.columns) {
        colStartOffsets.push_back(currentOffset);
        currentOffset += (col.size * numRows);
    }

    // Read rows by jumping between column blocks
    for (uint32 i = 0; i < numRows; i++) {
        std::vector<AppCUI::Utils::String> row;

        for (size_t c = 0; c < def.columns.size(); c++) {
            const auto& col = def.columns[c];
            // Offset = Start of Block + (Row Index * Item Size)
            uint32 valOffset = colStartOffsets[c] + (i * col.size);

            if (valOffset + col.size > buf.GetLength()) {
                row.emplace_back("<Corrupt>");
                continue;
            }

            if (col.type & MSICOL_INTEGER) {
                uint32 val = 0;
                if (col.size == 2) {
                    val = *(uint16*) (ptr + valOffset);
                    // Mask high bit (MSI internal flag)
                    val &= 0x7FFF;
                } else {
                    val = *(uint32*) (ptr + valOffset);
                    val &= 0x7FFFFFFF;
                }

                row.emplace_back(std::to_string(val).c_str());
            } else {
                uint32 strIdx = 0;
                if (stringBytes == 2)
                    strIdx = *(uint16*) (ptr + valOffset);
                else
                    strIdx = ptr[valOffset] | (ptr[valOffset + 1] << 8) | (ptr[valOffset + 2] << 16);
                row.emplace_back(GetString(strIdx).c_str());
            }
        }
        results.push_back(row);
    }
    return results;
}

const MsiTableDef* MSIFile::GetTableDefinition(const std::string& tableName) const
{
    auto it = tableDefs.find(tableName);
    if (it != tableDefs.end())
        return &it->second;
    return nullptr;
}

std::string MSIFile::ParseLpstr(const uint8_t* ptr, size_t avail)
{
    if (avail < 8)
        return "";

    uint32_t stringLen = 0;
    if (!read_u32_le(ptr + 4, avail - 4, stringLen))
        return "";

    if (stringLen == 0)
        return "";

    if (stringLen > (avail - 8))
        stringLen = static_cast<uint32_t>(avail - 8);

    std::string s(reinterpret_cast<const char*>(ptr + 8), stringLen);

    while (!s.empty() && s.back() == '\0') {
        s.pop_back();
    }

    return s;
}

