using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

internal static class Program
{
    private static dynamic? excel;
    private static string controlDirectory = "";
    private static readonly Dictionary<IntPtr, string> ownedWorkbooks = [];

    // All COM objects used by an operation stay alive until that operation ends.
    // The Application RCW alone belongs to Main and is released after Quit returns.
    // Indexed collection access avoids hidden COM enumerator RCWs from foreach.
    private sealed class ComScope : IDisposable
    {
        private readonly List<object> objects = [];
        private readonly HashSet<object> identities = new(ReferenceEqualityComparer.Instance);

        public dynamic Own(object? value)
        {
            if (value is not null && Marshal.IsComObject(value) &&
                !ReferenceEquals(value, (object?)excel) && identities.Add(value))
                objects.Add(value);
            return value!;
        }

        public void Dispose()
        {
            for (int index = objects.Count - 1; index >= 0; index--)
            {
                object value = objects[index];
                // Resolution may have promoted this same RCW to the root Application.
                if (!ReferenceEquals(value, (object?)excel)) Marshal.FinalReleaseComObject(value);
            }
            objects.Clear();
            identities.Clear();
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length < 2) return 2;
        string workbook = Path.GetFullPath(args[0]);
        controlDirectory = Path.GetFullPath(args[1]);
        Directory.CreateDirectory(controlDirectory);
        var start = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") { UseShellExecute = false };
        start.ArgumentList.Add("/x");
        start.ArgumentList.Add(workbook);
        using var process = Process.Start(start)!;
        try
        {
            var clock = Stopwatch.StartNew();
            while (excel is null && clock.Elapsed < TimeSpan.FromSeconds(45))
            {
                EnumWindows((window, _) =>
                {
                    GetWindowThreadProcessId(window, out uint pid);
                    if (pid != process.Id) return true;
                    EnumChildWindows(window, (child, _) =>
                    {
                        var name = new StringBuilder(128);
                        GetClassName(child, name, name.Capacity);
                        if (name.ToString() != "EXCEL7") return true;
                        var iid = new Guid("00020400-0000-0000-C000-000000000046");
                        if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref iid, out var native) == 0)
                        {
                            using var scope = new ComScope();
                            dynamic windowObject = scope.Own(native);
                            try { excel = windowObject.Application; } catch (COMException) { }
                        }
                        return excel is null;
                    }, IntPtr.Zero);
                    return excel is null;
                }, IntPtr.Zero);
                Application.DoEvents();
                Thread.Sleep(100);
            }
            if (excel is null) throw new InvalidOperationException("Own normal-start Excel instance could not be resolved.");
            // Never set COMAddIns.Connect, AutomationSecurity, macro policy or VBProject trust.
            RegisterInitialWorkbook(workbook);
            WriteState("started", process.Id);
            long sequence = 0;
            while (!process.HasExited)
            {
                Application.DoEvents();
                var commandPath = Path.Combine(controlDirectory, "command.json");
                if (File.Exists(commandPath))
                {
                    using var command = JsonDocument.Parse(File.ReadAllText(commandPath));
                    long id = command.RootElement.GetProperty("id").GetInt64();
                    if (id > sequence)
                    {
                        sequence = id;
                        string operation = command.RootElement.GetProperty("operation").GetString()!;
                        if (operation == "close") return CloseAndVerifyExit(process, sequence);
                        try { Execute(operation, command.RootElement); WriteState(operation, process.Id, sequence); }
                        catch (Exception error) { File.WriteAllText(Path.Combine(controlDirectory, $"result-{sequence}.json"), JsonSerializer.Serialize(new { operation, error = error.ToString() })); }
                    }
                }
                Thread.Sleep(100);
            }
            return 0;
        }
        catch (Exception error) { File.WriteAllText(Path.Combine(controlDirectory, "error.json"), error.ToString()); return 1; }
        finally
        {
            if (excel is not null)
            {
                try { CloseWorkbooksAndQuit(); }
                catch (Exception error) { File.WriteAllText(Path.Combine(controlDirectory, "cleanup-error.json"), error.ToString()); }
                finally { ReleaseApplication(); }
            }
            // DoEvents creates a WinForms thread context; release this harness's own context.
            Application.ExitThread();
        }
    }

    private sealed class UnknownWorkbooksException(IReadOnlyList<string> workbooks)
        : InvalidOperationException("Refusing to close workbooks or quit Excel because this process contains workbooks not owned by the harness: " + string.Join("; ", workbooks))
    {
        public IReadOnlyList<string> Workbooks { get; } = workbooks;
    }

    private static IntPtr WorkbookIdentity(object workbook)
    {
        IntPtr identity = Marshal.GetIUnknownForObject(workbook);
        try { return identity; }
        finally { Marshal.Release(identity); }
    }

    private static bool IsOwnedWorkbook(object workbook)
    {
        IntPtr identity = Marshal.GetIUnknownForObject(workbook);
        try { return ownedWorkbooks.ContainsKey(identity); }
        finally { Marshal.Release(identity); }
    }

    private static void RegisterOwnedWorkbook(object workbook)
    {
        // Hold one IUnknown reference so its identity cannot be recycled while tracked.
        IntPtr identity = Marshal.GetIUnknownForObject(workbook);
        string path;
        try { path = (string)((dynamic)workbook).FullName; }
        catch { Marshal.Release(identity); throw; }
        if (ownedWorkbooks.ContainsKey(identity)) Marshal.Release(identity);
        ownedWorkbooks[identity] = path;
    }

    private static void ReleaseOwnedWorkbookIdentities()
    {
        foreach (IntPtr identity in ownedWorkbooks.Keys) Marshal.Release(identity);
        ownedWorkbooks.Clear();
    }

    private static List<object> OpenWorkbooks(ComScope scope)
    {
        dynamic app = excel!;
        dynamic workbooks = scope.Own(app.Workbooks);
        var result = new List<object>();
        for (int index = 1; index <= (int)workbooks.Count; index++)
            result.Add((object)scope.Own(workbooks[index]));
        return result;
    }

    private static List<object> RequireAllWorkbooksOwned(ComScope scope)
    {
        List<object> workbooks = OpenWorkbooks(scope);
        var unknown = new List<string>();
        foreach (object workbook in workbooks)
            if (!IsOwnedWorkbook(workbook)) unknown.Add((string)((dynamic)workbook).FullName);
        if (unknown.Count != 0) throw new UnknownWorkbooksException(unknown);
        return workbooks;
    }

    private static void RegisterInitialWorkbook(string requestedPath)
    {
        using var scope = new ComScope();
        object? fixture = null;
        foreach (object workbook in OpenWorkbooks(scope))
        {
            string fullName = (string)((dynamic)workbook).FullName;
            if (!string.Equals(Path.GetFullPath(fullName), requestedPath, StringComparison.OrdinalIgnoreCase)) continue;
            if (fixture is not null) throw new InvalidOperationException("More than one initial fixture workbook matched.");
            fixture = workbook;
        }
        if (fixture is null) throw new InvalidOperationException("The requested initial fixture is not open in our Excel process.");
        RegisterOwnedWorkbook(fixture);
    }

    private static void RequireOwnedWorkbook(object workbook)
    {
        if (!IsOwnedWorkbook(workbook))
            throw new UnknownWorkbooksException([(string)((dynamic)workbook).FullName]);
    }

    private static void RejectUnownedExistingPath(ComScope scope, string requestedPath)
    {
        string fullPath = Path.GetFullPath(requestedPath);
        foreach (object workbook in OpenWorkbooks(scope))
            if (string.Equals(Path.GetFullPath((string)((dynamic)workbook).FullName), fullPath, StringComparison.OrdinalIgnoreCase))
                RequireOwnedWorkbook(workbook);
    }

    private static void CloseWorkbooksAndQuit()
    {
        dynamic app = excel!;
        using (var scope = new ComScope())
        {
            // Inspect everything before closing anything. Other automation can attach even to /x.
            List<object> workbooks = RequireAllWorkbooksOwned(scope);
            foreach (object workbook in workbooks)
            {
                // Recheck between operations; close only the recorded object, never Workbooks[1].
                RequireAllWorkbooksOwned(scope);
                RequireOwnedWorkbook(workbook);
                ((dynamic)workbook).Close(false);
            }
            if (RequireAllWorkbooksOwned(scope).Count != 0)
                throw new InvalidOperationException("An owned workbook remained open; Excel Quit was not requested.");
        }
        ReleaseOwnedWorkbookIdentities();
        // All workbooks must still be absent immediately before Quit.
        using (var scope = new ComScope())
            if (OpenWorkbooks(scope).Count != 0)
                throw new UnknownWorkbooksException(OpenWorkbooks(scope).Select(value => (string)((dynamic)value).FullName).ToArray());
        app.Quit();
    }

    private static void ReleaseApplication()
    {
        ReleaseOwnedWorkbookIdentities();
        object? application = (object?)excel;
        excel = null;
        if (application is not null && Marshal.IsComObject(application))
            Marshal.FinalReleaseComObject(application);
    }

    private static int CloseAndVerifyExit(Process process, long sequence)
    {
        try { CloseWorkbooksAndQuit(); }
        catch (UnknownWorkbooksException error)
        {
            // Release our references only; leave unknown documents and this Excel process running.
            ReleaseApplication();
            File.WriteAllText(Path.Combine(controlDirectory, "close-refused.json"),
                JsonSerializer.Serialize(new { processId = process.Id, sequence, quitRequested = false,
                    reason = "unknown-workbook", unknownWorkbooks = error.Workbooks, error = error.Message }));
            return 4;
        }
        ReleaseApplication();
        File.WriteAllText(Path.Combine(controlDirectory, "quit-requested.json"),
            JsonSerializer.Serialize(new { processId = process.Id, sequence, quitReturned = true, comReferencesReleased = true }));

        // Pump this STA while waiting for the exact process we started. Never force termination.
        var elapsed = Stopwatch.StartNew();
        while (!process.HasExited && elapsed.Elapsed < TimeSpan.FromSeconds(20))
        {
            Application.DoEvents();
            Thread.Sleep(100);
        }
        bool excelExited = process.HasExited;
        File.WriteAllText(Path.Combine(controlDirectory, "closed.json"),
            JsonSerializer.Serialize(new { processId = process.Id, sequence, quitReturned = true,
                comReferencesReleased = true, excelExited, waitMilliseconds = elapsed.ElapsedMilliseconds,
                exitCode = excelExited ? (int?)process.ExitCode : null }));
        return excelExited ? 0 : 3;
    }

    private static void Execute(string operation, JsonElement command)
    {
        using var scope = new ComScope();
        dynamic app = excel!;
        dynamic sheet = scope.Own(app.ActiveSheet);
        if (operation is not ("open" or "ordinary" or "activate" or "snapshot"))
        {
            dynamic activeWorkbook = scope.Own(app.ActiveWorkbook);
            RequireOwnedWorkbook((object)activeWorkbook);
        }
        switch (operation)
        {
            case "select":
            {
                dynamic range = scope.Own(sheet.Range[command.GetProperty("range").GetString()!]);
                range.Select();
                break;
            }
            case "filter7":
            {
                dynamic tables = scope.Own(sheet.ListObjects);
                dynamic table = scope.Own(tables[1]);
                dynamic tableRange = scope.Own(table.Range);
                tableRange.AutoFilter(1, "<=f007.txt");
                dynamic range = scope.Own(sheet.Range["A2:A101"]);
                range.Select();
                break;
            }
            case "hide":
            {
                dynamic rows = scope.Own(sheet.Rows);
                dynamic row = scope.Own(rows[command.GetProperty("row").GetInt32()]);
                row.Hidden = true;
                break;
            }
            case "unfilter":
            {
                if ((bool)sheet.FilterMode) sheet.ShowAllData();
                dynamic rows = scope.Own(sheet.Rows);
                rows.Hidden = false;
                break;
            }
            case "sort-desc":
            {
                dynamic tables = scope.Own(sheet.ListObjects);
                dynamic table = scope.Own(tables[1]);
                dynamic sort = scope.Own(table.Sort);
                dynamic fields = scope.Own(sort.SortFields);
                dynamic columns = scope.Own(table.ListColumns);
                dynamic column = scope.Own(columns["이름"]);
                dynamic data = scope.Own(column.DataBodyRange);
                fields.Clear();
                scope.Own(fields.Add(data, 0, 2));
                sort.Header = 1;
                sort.Apply();
                break;
            }
            case "sheet":
            {
                dynamic workbook = scope.Own(app.ActiveWorkbook);
                dynamic sheets = scope.Own(workbook.Worksheets);
                dynamic selectedSheet = scope.Own(sheets[command.GetProperty("name").GetString()!]);
                selectedSheet.Activate();
                dynamic range = scope.Own(selectedSheet.Range[command.GetProperty("range").GetString()!]);
                range.Select();
                break;
            }
            case "snapshot":
                break;
            case "activate":
            {
                dynamic workbooks = scope.Own(app.Workbooks);
                dynamic workbook = scope.Own(workbooks[command.GetProperty("name").GetString()!]);
                RequireOwnedWorkbook((object)workbook);
                workbook.Activate();
                break;
            }
            case "ordinary":
            {
                dynamic workbooks = scope.Own(app.Workbooks);
                dynamic ordinary = scope.Own(workbooks.Add());
                RegisterOwnedWorkbook((object)ordinary);
                dynamic ordinarySheet = scope.Own(ordinary.ActiveSheet);
                dynamic a1 = scope.Own(ordinarySheet.Range["A1"]);
                dynamic b1 = scope.Own(ordinarySheet.Range["B1"]);
                dynamic a2 = scope.Own(ordinarySheet.Range["A2"]);
                dynamic b2 = scope.Own(ordinarySheet.Range["B2"]);
                a1.Value2 = "이름";
                b1.Value2 = "전체경로";
                a2.Value2 = "f001.txt";
                b2.Value2 = @"C:\fake.txt";
                dynamic tables = scope.Own(ordinarySheet.ListObjects);
                dynamic range = scope.Own(ordinarySheet.Range["A1:B2"]);
                scope.Own(tables.Add(1, range, Type.Missing, 1));
                a2.Select();
                break;
            }
            case "drop-technical":
            {
                dynamic tables = scope.Own(sheet.ListObjects);
                dynamic table = scope.Own(tables[1]);
                dynamic columns = scope.Own(table.ListColumns);
                dynamic column = scope.Own(columns["__FLT_SourceRecord"]);
                column.Delete();
                dynamic range = scope.Own(sheet.Range["A3"]);
                range.Select();
                break;
            }
            case "reorder-path":
            {
                dynamic tables = scope.Own(sheet.ListObjects);
                dynamic table = scope.Own(tables[1]);
                dynamic columns = scope.Own(table.ListColumns);
                dynamic column = scope.Own(columns["전체경로"]);
                dynamic range = scope.Own(column.Range);
                dynamic entireColumn = scope.Own(range.EntireColumn);
                entireColumn.Cut();
                dynamic sheetColumns = scope.Own(sheet.Columns);
                dynamic firstColumn = scope.Own(sheetColumns["A:A"]);
                firstColumn.Insert(-4161);
                app.CutCopyMode = false;
                break;
            }
            case "open":
            {
                dynamic workbooks = scope.Own(app.Workbooks);
                string path = command.GetProperty("path").GetString()!;
                RejectUnownedExistingPath(scope, path);
                // Keep the pre-existing RCWs in this scope so their identities remain stable.
                var existing = OpenWorkbooks(scope).Select(WorkbookIdentity).ToHashSet();
                dynamic opened = scope.Own(workbooks.Open(path));
                // Excel can return an existing workbook for a path alias; never claim that object.
                if (existing.Contains(WorkbookIdentity((object)opened))) RequireOwnedWorkbook((object)opened);
                RegisterOwnedWorkbook((object)opened);
                break;
            }
            case "rename-sheet":
                sheet.Name = "검증_이름변경";
                break;
            case "save-as":
            {
                dynamic workbook = scope.Own(app.ActiveWorkbook);
                workbook.SaveAs(command.GetProperty("path").GetString()!, 51);
                RegisterOwnedWorkbook((object)workbook);
                break;
            }
            case "execute":
            {
                dynamic button = Button(scope);
                button.Execute();
                break;
            }
            case "set-schema":
            {
                dynamic workbook = scope.Own(app.ActiveWorkbook);
                dynamic properties = scope.Own(workbook.CustomDocumentProperties);
                dynamic property = scope.Own(properties["FLT.ActionSchemaVersion"]);
                property.Value = command.GetProperty("value").GetString();
                dynamic range = scope.Own(sheet.Range["A2"]);
                range.Select();
                break;
            }
            case "set-path":
            {
                dynamic tables = scope.Own(sheet.ListObjects);
                dynamic table = scope.Own(tables[1]);
                dynamic columns = scope.Own(table.ListColumns);
                dynamic column = scope.Own(columns["전체경로"]);
                dynamic data = scope.Own(column.DataBodyRange);
                dynamic cells = scope.Own(data.Cells);
                dynamic cell = scope.Own(cells[1, 1]);
                cell.Value2 = command.GetProperty("value").GetString();
                dynamic range = scope.Own(sheet.Range["A2"]);
                range.Select();
                break;
            }
            default: throw new ArgumentException("Unsupported acceptance command.");
        }
        Application.DoEvents();
    }

    private static dynamic Button(ComScope scope)
    {
        dynamic app = excel!;
        dynamic bars = scope.Own(app.CommandBars);
        dynamic cellBar = scope.Own(bars["Cell"]);
        return scope.Own(cellBar.FindControl(Type.Missing, Type.Missing, "FileListToExcel.ExcelAddIn.Collect.v1", Type.Missing, true));
    }

    private static void WriteState(string operation, int pid, long sequence = 0)
    {
        using var scope = new ComScope();
        dynamic app = excel!;
        bool connected = false;
        string? addinError = null;
        try
        {
            dynamic addins = scope.Own(app.COMAddIns);
            dynamic addin = scope.Own(addins["FileListToExcel.ExcelAddIn"]);
            connected = addin.Connect;
        }
        catch (Exception error) { addinError = error.Message; }
        dynamic? button = Button(scope);
        var menuCounts = new Dictionary<string, int>();
        dynamic bars = scope.Own(app.CommandBars);
        foreach (var barName in new[] { "Cell", "List Range Popup" })
        {
            int count = 0;
            dynamic bar = scope.Own(bars[barName]);
            dynamic controls = scope.Own(bar.Controls);
            for (int index = 1; index <= (int)controls.Count; index++)
            {
                dynamic control = scope.Own(controls[index]);
                if ((string)control.Tag == "FileListToExcel.ExcelAddIn.Menu.v1") count++;
            }
            menuCounts[barName] = count;
        }
        var props = new Dictionary<string, string>();
        dynamic workbook = scope.Own(app.ActiveWorkbook);
        dynamic properties = scope.Own(workbook.CustomDocumentProperties);
        for (int index = 1; index <= (int)properties.Count; index++)
        {
            dynamic property = scope.Own(properties[index]);
            props[(string)property.Name] = Convert.ToString(property.Value);
        }
        dynamic selection = scope.Own(app.Selection);
        File.WriteAllText(Path.Combine(controlDirectory, sequence == 0 ? "started.json" : $"result-{sequence}.json"),
            JsonSerializer.Serialize(new { operation, processId = pid, version = (string)app.Version, build = Convert.ToString(app.Build),
                eventsEnabled = (bool)app.EnableEvents, windowHandle = (long)app.Hwnd, menuCounts, connected, addinError, menuPresent = button is not null, enabled = button is not null && (bool)button.Enabled,
                workbook = (string)workbook.FullName, selection = (string)selection.Address, saved = (bool)workbook.Saved, properties = props },
                new JsonSerializerOptions { WriteIndented = true }));
    }

    private delegate bool EnumWindowProc(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object native);
}
