#pragma once
#include "../FileListToExcel.Shell/RequestCodec.h"
#include "Automation.h"
#include <set>
namespace flt
{
struct SelectedTable
{
    Ptr<IDispatch> workbook, table, body, selection;
    std::wstring workbookId, tableName;
};
inline Ptr<IDispatch> Intersection(IDispatch *app, IDispatch *a, IDispatch *b)
{
    auto v = Call(app, L"Intersect", DISPATCH_METHOD, {V(a), V(b)});
    return v.vt == VT_DISPATCH ? Ptr<IDispatch>(v.pdispVal, true) : Ptr<IDispatch>();
}
inline SelectedTable Locate(IDispatch *app)
{
    SelectedTable result;
    result.workbook = Obj(app, L"ActiveWorkbook");
    if (Property(result.workbook.get(), L"FLT.Product") != L"FileListToExcel" ||
        Property(result.workbook.get(), L"FLT.ActionSchemaVersion") != L"1")
        throw std::runtime_error("Unsupported result version");
    result.workbookId = Property(result.workbook.get(), L"FLT.WorkbookId");
    if (!IsGuid(result.workbookId))
        throw std::runtime_error("Invalid workbook identifier");
    result.selection = Obj(app, L"Selection");
    auto sheet = Obj(app, L"ActiveSheet");
    auto tables = Obj(sheet.get(), L"ListObjects");
    long count = Call(tables.get(), L"Count").Number();
    if (count > 1024)
        throw std::runtime_error("Too many tables");
    for (long n = 1; n <= count; ++n)
    {
        auto table = Obj(tables.get(), L"Item", {V(n)});
        auto range = Obj(table.get(), L"Range");
        if (!Intersection(app, result.selection.get(), range.get()))
            continue;
        if (result.table)
            throw std::runtime_error("Select rows from one table");
        auto name = Call(table.get(), L"Name").Text();
        if (name.size() <= 13 || name.size() > 32 || name.substr(0, 13) != L"FileListTable" ||
            name[13] < L'1' || name[13] > L'9' ||
            name.find_first_not_of(L"0123456789", 13) != std::wstring::npos)
            throw std::runtime_error("Unsupported table name");
        auto role = Property(result.workbook.get(), L"FLT.Table." + name);
        if (role != L"files" && role != L"duplicates" && role != L"matches")
            throw std::runtime_error("Unsupported table");
        result.table = table;
        result.tableName = name;
    }
    if (!result.table)
        throw std::runtime_error("Select a File List result table");
    result.body = Obj(result.table.get(), L"DataBodyRange");
    auto columns = Obj(result.table.get(), L"ListColumns");
    for (auto name : {L"__FLT_ItemId", L"__FLT_SourceRecord", L"전체경로", L"이름"})
        Obj(columns.get(), L"Item", {V(name)});
    return result;
}
inline std::wstring NewId()
{
    GUID id{};
    Check(CoCreateGuid(&id));
    wchar_t text[40]{};
    StringFromGUID2(id, text, 40);
    std::wstring value(text + 1, 36);
    for (auto &c : value)
        if (c >= L'A' && c <= L'F')
            c = static_cast<wchar_t>(c + 32);
    return value;
}
inline std::string Capture(IDispatch *app, const std::wstring &requestId)
{
    auto selected = Locate(app);
    auto areas = Obj(selected.selection.get(), L"Areas");
    long count = Call(areas.get(), L"Count").Number();
    if (count > 10000)
        throw std::runtime_error("Too many selection areas");
    std::set<long> rows;
    for (long n = 1; n <= count; ++n)
    {
        auto area = Obj(areas.get(), L"Item", {V(n)});
        auto hit = Intersection(app, area.get(), selected.body.get());
        if (!hit)
            throw std::runtime_error("Selection includes unrelated cells or only headers");
        auto areaCols = Obj(area.get(), L"Columns");
        auto tableRange = Obj(selected.table.get(), L"Range");
        auto withinTable = Intersection(app, area.get(), tableRange.get());
        auto hitCols = Obj(withinTable.get(), L"Columns");
        auto areaRows = Obj(area.get(), L"Rows");
        auto hitRows = Obj(withinTable.get(), L"Rows");
        bool wholeRows = Call(areaCols.get(), L"Count").Number() == 16384;
        if (!wholeRows &&
            (Call(areaCols.get(), L"Count").Number() != Call(hitCols.get(), L"Count").Number() ||
             Call(areaRows.get(), L"Count").Number() != Call(hitRows.get(), L"Count").Number()))
            throw std::runtime_error("Selection includes unrelated cells");
        auto entireRows = Obj(hit.get(), L"EntireRow");
        auto rowHit = Intersection(app, entireRows.get(), selected.body.get());
        Ptr<IDispatch> visible;
        try
        {
            visible = Call(rowHit.get(), L"SpecialCells", DISPATCH_METHOD, {V(12L)}).Obj();
        }
        catch (...)
        {
            auto allRows = Obj(rowHit.get(), L"EntireRow");
            auto hidden = Call(allRows.get(), L"Hidden");
            if (hidden.vt == VT_BOOL && hidden.boolVal == VARIANT_TRUE)
                continue;
            throw;
        }
        auto visibleAreas = Obj(visible.get(), L"Areas");
        long visibleCount = Call(visibleAreas.get(), L"Count").Number();
        for (long i = 1; i <= visibleCount; ++i)
        {
            auto visibleArea = Obj(visibleAreas.get(), L"Item", {V(i)});
            long start = Call(visibleArea.get(), L"Row").Number();
            auto vr = Obj(visibleArea.get(), L"Rows");
            long length = Call(vr.get(), L"Count").Number();
            for (long r = start; r < start + length; ++r)
            {
                rows.insert(r);
                if (rows.size() > 10000)
                    throw std::runtime_error("Select at most 10000 visible rows");
            }
        }
    }
    if (rows.empty())
        throw std::runtime_error("No visible data rows");
    auto columns = Obj(selected.table.get(), L"ListColumns");
    auto columnIndex = [&](const wchar_t *name) {
        auto c = Obj(columns.get(), L"Item", {V(name)});
        return Call(c.get(), L"Index").Number();
    };
    long idCol = columnIndex(L"__FLT_ItemId"), sourceCol = columnIndex(L"__FLT_SourceRecord"),
         pathCol = columnIndex(L"전체경로"), nameCol = columnIndex(L"이름");
    long firstRow = Call(selected.body.get(), L"Row").Number();
    auto cells = Obj(selected.body.get(), L"Cells");
    std::string payload =
        "{\"mode\":\"collect-files\",\"version\":1,\"requestId\":" + filelist::JsonString(requestId) +
        ",\"workbookId\":" + filelist::JsonString(selected.workbookId) +
        ",\"tableName\":" + filelist::JsonString(selected.tableName) + ",\"rows\":[";
    std::set<std::wstring> ids;
    bool comma = false;
    for (long row : rows)
    {
        auto value = [&](long col) {
            auto cell = Obj(cells.get(), L"Item", {V(row - firstRow + 1), V(col)});
            if (Call(cell.get(), L"HasFormula").Truth())
                throw std::runtime_error("Formula in source metadata");
            return Call(cell.get(), L"Value2").Text();
        };
        auto itemId = value(idCol), record = value(sourceCol), path = value(pathCol), name = value(nameCol);
        if (!IsGuid(itemId) || !ids.insert(itemId).second || record.empty() || record.size() > 32767 ||
            record.front() != L'{' || record.back() != L'}')
            throw std::runtime_error("Invalid or duplicated source record");
        // The helper parses bounded JSON and checks displayed name/path, including unavailable rows.
        if (comma)
            payload += ',';
        comma = true;
        payload += "{\"itemId\":" + filelist::JsonString(itemId) +
                   ",\"displayPath\":" + filelist::JsonString(path) +
                   ",\"displayName\":" + filelist::JsonString(name) +
                   ",\"sourceRecord\":" + filelist::JsonString(record) + "}";
        if (payload.size() > filelist::MaximumRequestBytes - 2)
            throw std::runtime_error("Selection request exceeds 32 MiB");
    }
    return payload + "]}";
}
} // namespace flt
