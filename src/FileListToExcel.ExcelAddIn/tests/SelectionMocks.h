#pragma once
#include "SdiMocks.h"
#include <set>

namespace flt::test
{
struct Rect
{
    long top, bottom, left, right;
};

// Values are returned through IDispatch so these regressions execute the same
// Automation/Locate/Capture path used by the real Excel click callback.
struct SelectionFixture
{
    std::vector<Ptr<Node>> nodes;
    std::vector<Rect> selection{{2, 2, 1, 1}};
    std::set<long> hiddenRows;
    long lastRow = 101;
    bool missingSourceColumn = false;
    long formulaRow = 0, duplicateIdRow = 0;
    unsigned cellReads = 0;
    Ptr<Node> app, workbook, table;

    Ptr<Node> Make(std::function<V(const std::wstring &, WORD, DISPPARAMS *)> action)
    {
        Ptr<Node> node(new Node());
        node->action = std::move(action);
        nodes.push_back(node);
        return node;
    }
    static long Arg(DISPPARAMS *args, size_t index)
    {
        return args->rgvarg[args->cArgs - index - 1].lVal;
    }
    static std::wstring Id(long row)
    {
        wchar_t text[40]{};
        swprintf_s(text, L"aaaaaaaa-0000-0000-0000-%012ld", row);
        return text;
    }
    std::map<IDispatch *, std::vector<Rect>> ranges;
    Ptr<Node> Range(std::vector<Rect> rectangles)
    {
        auto node = Make([this, rectangles](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name == L"Row") return V(rectangles.at(0).top);
            if (name == L"Rows" || name == L"Columns")
            {
                long count = name == L"Rows" ? rectangles.at(0).bottom - rectangles.at(0).top + 1
                                             : rectangles.at(0).right - rectangles.at(0).left + 1;
                return V(Make([count](const std::wstring &field, WORD, DISPPARAMS *) -> V {
                    if (field == L"Count") return V(count);
                    throw std::runtime_error("Unexpected dimension member");
                }).get());
            }
            if (name == L"Areas")
                return V(Make([this, rectangles](const std::wstring &field, WORD, DISPPARAMS *parameters) -> V {
                    if (field == L"Count") return V(static_cast<long>(rectangles.size()));
                    if (field == L"Item") return V(Range({rectangles.at(Arg(parameters, 0) - 1)}).get());
                    throw std::runtime_error("Unexpected area member");
                }).get());
            if (name == L"EntireRow")
            {
                auto rows = rectangles;
                for (auto &r : rows) { r.left = 1; r.right = 16384; }
                return V(Range(std::move(rows)).get());
            }
            if (name == L"Hidden")
            {
                bool allHidden = true;
                for (const auto &r : rectangles)
                    for (long row = r.top; row <= r.bottom; ++row)
                        allHidden = allHidden && hiddenRows.count(row) != 0;
                return V(allHidden);
            }
            if (name == L"SpecialCells")
            {
                if (Arg(args, 0) != 12) throw std::runtime_error("Expected visible-cell filter");
                std::vector<Rect> visible;
                for (const auto &r : rectangles)
                    for (long row = r.top; row <= r.bottom; ++row)
                        if (!hiddenRows.count(row))
                            visible.push_back({row, row, r.left, r.right});
                if (visible.empty()) throw std::runtime_error("No visible cells");
                return V(Range(std::move(visible)).get());
            }
            if (name == L"Cells")
                return V(Make([this](const std::wstring &field, WORD, DISPPARAMS *parameters) -> V {
                    if (field != L"Item") throw std::runtime_error("Expected cell index");
                    long row = Arg(parameters, 0) + 1, column = Arg(parameters, 1);
                    return V(Make([this, row, column](const std::wstring &member, WORD, DISPPARAMS *) -> V {
                        if (member == L"HasFormula") return V(row == formulaRow);
                        if (member != L"Value2") throw std::runtime_error("Unexpected cell member");
                        ++cellReads;
                        if (column == 1) return V((L"file" + std::to_wstring(row) + L".txt"));
                        if (column == 2) return V((L"C:\\fixture\\file" + std::to_wstring(row) + L".txt"));
                        if (column == 3) return V(Id(row == duplicateIdRow ? 2 : row));
                        if (column == 4) return V(L"{\"kind\":\"file\"}");
                        throw std::runtime_error("Unexpected metadata column");
                    }).get());
                }).get());
            throw std::runtime_error("Unexpected range member");
        });
        ranges[node.get()] = std::move(rectangles);
        return node;
    }
    SelectionFixture()
    {
        auto properties = Make([this](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name != L"Item") throw std::runtime_error("Expected custom property");
            std::wstring key(args->rgvarg[0].bstrVal), value;
            if (key == L"FLT.Product") value = L"FileListToExcel";
            else if (key == L"FLT.ActionSchemaVersion") value = L"1";
            else if (key == L"FLT.WorkbookId") value = Id(1);
            else if (key == L"FLT.Table.FileListTable1") value = L"files";
            else throw std::runtime_error("Unknown custom property");
            return V(Make([value](const std::wstring &field, WORD, DISPPARAMS *) -> V {
                if (field == L"Value") return V(value);
                throw std::runtime_error("Expected property value");
            }).get());
        });
        workbook = Make([properties](const std::wstring &name, WORD, DISPPARAMS *) -> V {
            if (name == L"CustomDocumentProperties") return V(properties.get());
            throw std::runtime_error("Unexpected workbook member");
        });
        auto columns = Make([this](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name != L"Item") throw std::runtime_error("Expected named column");
            std::wstring column(args->rgvarg[0].bstrVal);
            long index = column == L"이름" ? 1 : column == L"전체경로" ? 2 :
                         column == L"__FLT_ItemId" ? 3 : column == L"__FLT_SourceRecord" ? 4 : 0;
            if (!index || (index == 4 && missingSourceColumn)) throw std::runtime_error("Missing column");
            return V(Make([index](const std::wstring &field, WORD, DISPPARAMS *) -> V {
                if (field == L"Index") return V(index);
                throw std::runtime_error("Expected column index");
            }).get());
        });
        table = Make([this, columns](const std::wstring &name, WORD, DISPPARAMS *) -> V {
            if (name == L"Name") return V(L"FileListTable1");
            if (name == L"Range") return V(Range({{1, lastRow, 1, 4}}).get());
            if (name == L"DataBodyRange") return V(Range({{2, lastRow, 1, 4}}).get());
            if (name == L"ListColumns") return V(columns.get());
            throw std::runtime_error("Unexpected table member");
        });
        auto tables = Make([this](const std::wstring &name, WORD, DISPPARAMS *) -> V {
            if (name == L"Count") return V(1L);
            if (name == L"Item") return V(table.get());
            throw std::runtime_error("Unexpected tables member");
        });
        auto sheet = Make([tables](const std::wstring &name, WORD, DISPPARAMS *) -> V {
            if (name == L"ListObjects") return V(tables.get());
            throw std::runtime_error("Unexpected sheet member");
        });
        app = Make([this, sheet](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name == L"ActiveWorkbook") return V(workbook.get());
            if (name == L"Selection") return V(Range(selection).get());
            if (name == L"ActiveSheet") return V(sheet.get());
            if (name == L"Intersect")
            {
                std::vector<Rect> intersections;
                for (const auto &a : ranges.at(args->rgvarg[1].pdispVal))
                    for (const auto &b : ranges.at(args->rgvarg[0].pdispVal))
                    {
                        Rect r{std::max(a.top, b.top), std::min(a.bottom, b.bottom),
                               std::max(a.left, b.left), std::min(a.right, b.right)};
                        if (r.top <= r.bottom && r.left <= r.right) intersections.push_back(r);
                    }
                return intersections.empty() ? V() : V(Range(std::move(intersections)).get());
            }
            throw std::runtime_error("Unexpected application member");
        });
    }
};
inline size_t Occurrences(const std::string &text, const std::string &token)
{
    size_t count = 0, position = 0;
    while ((position = text.find(token, position)) != std::string::npos) { ++count; position += token.size(); }
    return count;
}
inline void SelectionRegression()
{
    auto require = [](bool value, const char *message) { if (!value) throw std::runtime_error(message); };
    auto capture = [](SelectionFixture &fixture) { return Capture(fixture.app.get(), SelectionFixture::Id(999)); };
    auto reject = [&](SelectionFixture &fixture, const char *expected, const char *message) {
        bool rejected = false;
        try { capture(fixture); }
        catch (const std::runtime_error &error) { rejected = std::string(error.what()) == expected; }
        require(rejected, message);
    };
    {
        SelectionFixture f;
        f.selection = {{5, 5, 1, 1}, {2, 3, 1, 1}, {2, 2, 2, 2}};
        auto payload = capture(f);
        require(Occurrences(payload, "\"itemId\"") == 3 && f.cellReads == 12,
                "Multi-area capture must deduplicate rows and read each selected row once");
        require(payload.find("file2.txt") < payload.find("file3.txt") && payload.find("file3.txt") < payload.find("file5.txt"),
                "Capture must preserve table row order independently of selection area order");
    }
    {
        SelectionFixture f;
        f.selection = {{2, 8, 1, 1}};
        f.hiddenRows = {3, 4, 6, 7};
        auto payload = capture(f);
        require(Occurrences(payload, "\"itemId\"") == 3 && payload.find("file3.txt") == std::string::npos && f.cellReads == 12,
                "Hidden and filtered rows must not be read or dispatched");
    }
    {
        SelectionFixture f;
        f.selection = {{1, 1048576, 1, 16384}};
        auto payload = capture(f);
        require(Occurrences(payload, "\"itemId\"") == 100 && f.cellReads == 400,
                "Whole-sheet selection must intersect the table before enumerating rows");
    }
    {
        SelectionFixture f;
        f.selection = {{2, 3, 1, 1}};
        f.hiddenRows = {2, 3};
        reject(f, "No visible data rows", "All-hidden selection must not dispatch");
        require(f.cellReads == 0, "All-hidden selection must not read source metadata");
    }
    {
        SelectionFixture f;
        f.selection = {{1, 1, 1, 4}};
        reject(f, "Selection includes unrelated cells or only headers", "Header-only selection must not dispatch");
    }
    {
        SelectionFixture f;
        f.selection = {{2, 2, 1, 1}, {200, 200, 1, 1}};
        reject(f, "Selection includes unrelated cells or only headers", "Unrelated selected area must reject the whole request");
        require(f.cellReads == 0, "Mixed selection must fail before metadata reads");
    }
    {
        SelectionFixture f;
        f.selection = {{2, 2, 1, 5}};
        reject(f, "Selection includes unrelated cells", "Partial rows extending outside the table must not dispatch");
    }
    {
        SelectionFixture f;
        f.missingSourceColumn = true;
        reject(f, "Excel automation call failed", "Deleted technical column must not dispatch");
    }
    {
        SelectionFixture f;
        f.formulaRow = 2;
        reject(f, "Formula in source metadata", "Formula source metadata must not dispatch");
    }
    {
        SelectionFixture f;
        f.selection = {{2, 3, 1, 1}};
        f.duplicateIdRow = 3;
        reject(f, "Invalid or duplicated source record", "Duplicate item identities must not dispatch");
    }
    {
        SelectionFixture f;
        f.lastRow = 10002;
        f.selection = {{2, 10002, 1, 1}};
        reject(f, "Select at most 10000 visible rows", "More than 10000 visible rows must not dispatch");
        require(f.cellReads == 0, "Row limit must precede source metadata reads");
    }
}
} // namespace flt::test
