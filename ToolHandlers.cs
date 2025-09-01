using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel_mcp_dotnet;

public static class ToolHandlers
{
    private static readonly ExcelHandler _excelHandler = new ExcelHandler();

    public static async Task<JObject> HandleWorkbookMetadata(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        if (string.IsNullOrEmpty(filepath)) throw new ArgumentException("filepath required");

        var metadata = await _excelHandler.GetWorkbookMetadataAsync(filepath);
        return JObject.FromObject(metadata);
    }

    public static async Task<JObject> HandleDataImportCsv(JObject parameters)
    {
        var csvPath = parameters["csv_path"]?.ToString();
        var excelPath = parameters["excel_path"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString() ?? "Sheet1";
        var hasHeader = parameters["has_header"]?.ToObject<bool>() ?? true;

        if (string.IsNullOrEmpty(csvPath) || string.IsNullOrEmpty(excelPath)) throw new ArgumentException("csv_path and excel_path required");

        await _excelHandler.ImportCsvAsync(csvPath, excelPath, sheetName, hasHeader);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleDataExportCsv(JObject parameters)
    {
        var excelPath = parameters["excel_path"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var csvPath = parameters["csv_path"]?.ToString();

        if (string.IsNullOrEmpty(excelPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(csvPath)) throw new ArgumentException("excel_path, sheet_name, and csv_path required");

        await _excelHandler.ExportCsvAsync(excelPath, sheetName, csvPath);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleFormatRange(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var startCell = parameters["start_cell"]?.ToString();
        var endCell = parameters["end_cell"]?.ToString();
        var bold = parameters["bold"]?.ToObject<bool>() ?? false;
        var fillColor = parameters["fill_color"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(startCell) || string.IsNullOrEmpty(endCell)) throw new ArgumentException("filepath, sheet_name, start_cell, and end_cell required");

        await _excelHandler.FormatRangeAsync(filepath, sheetName, startCell, endCell, bold, fillColor);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleFormulaApply(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var cell = parameters["cell"]?.ToString();
        var formula = parameters["formula"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cell) || string.IsNullOrEmpty(formula)) throw new ArgumentException("filepath, sheet_name, cell, and formula required");

        await _excelHandler.ApplyFormulaAsync(filepath, sheetName, cell, formula);
        return new JObject { ["success"] = true };
    }

    // Advanced formatting handlers
    public static async Task<JObject> HandleFormatAdvanced(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var formatting = parameters["formatting"]?.ToObject<AdvancedFormatting>();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range) || formatting == null)
            throw new ArgumentException("filepath, sheet_name, range, and formatting required");

        await _excelHandler.ApplyAdvancedFormattingAsync(filepath, sheetName, range, formatting);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleFormatConditional(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var ruleType = parameters["rule_type"]?.ToString();
        var condition = parameters["condition"]?.ToObject<ConditionalFormattingCondition>();
        var format = parameters["format"]?.ToObject<ConditionalFormattingFormat>();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range) ||
            string.IsNullOrEmpty(ruleType) || condition == null || format == null)
            throw new ArgumentException("All parameters required for conditional formatting");

        await _excelHandler.ApplyConditionalFormattingAsync(filepath, sheetName, range, ruleType, condition, format);
        return new JObject { ["success"] = true };
    }

    // Data manipulation handlers
    public static async Task<JObject> HandleDataSort(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var sortBy = parameters["sort_by"]?.ToObject<List<SortColumn>>();
        var ascending = parameters["ascending"]?.ToObject<bool>() ?? true;

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range) || sortBy == null)
            throw new ArgumentException("filepath, sheet_name, range, and sort_by required");

        await _excelHandler.SortRangeAsync(filepath, sheetName, range, sortBy, ascending);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleDataFilter(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var filters = parameters["filters"]?.ToObject<Dictionary<string, object>>();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range))
            throw new ArgumentException("filepath, sheet_name, and range required");

        await _excelHandler.ApplyFiltersAsync(filepath, sheetName, range, filters);
        return new JObject { ["success"] = true };
    }

    // Cell operations
    public static async Task<JObject> HandleCellMerge(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range))
            throw new ArgumentException("filepath, sheet_name, and range required");

        await _excelHandler.MergeCellsAsync(filepath, sheetName, range);
        return new JObject { ["success"] = true };
    }

    public static async Task<JObject> HandleCellUnmerge(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range))
            throw new ArgumentException("filepath, sheet_name, and range required");

        await _excelHandler.UnmergeCellsAsync(filepath, sheetName, range);
        return new JObject { ["success"] = true };
    }

    // Named ranges
    public static async Task<JObject> HandleNamedRangeCreate(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var name = parameters["name"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range))
            throw new ArgumentException("filepath, name, sheet_name, and range required");

        await _excelHandler.CreateNamedRangeAsync(filepath, name, sheetName, range);
        return new JObject { ["success"] = true };
    }

    // Charts
    public static async Task<JObject> HandleChartCreate(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var dataRange = parameters["data_range"]?.ToString();
        var chartType = parameters["chart_type"]?.ToString();
        var targetCell = parameters["target_cell"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(dataRange) ||
            string.IsNullOrEmpty(chartType) || string.IsNullOrEmpty(targetCell))
            throw new ArgumentException("All parameters required for chart creation");

        await _excelHandler.CreateChartAsync(filepath, sheetName, dataRange, chartType, targetCell);
        return new JObject { ["success"] = true };
    }

    // Pivot tables
    public static async Task<JObject> HandlePivotCreate(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sourceSheet = parameters["source_sheet"]?.ToString();
        var sourceRange = parameters["source_range"]?.ToString();
        var targetSheet = parameters["target_sheet"]?.ToString();
        var targetCell = parameters["target_cell"]?.ToString();
        var rows = parameters["rows"]?.ToObject<List<string>>();
        var columns = parameters["columns"]?.ToObject<List<string>>();
        var values = parameters["values"]?.ToObject<List<PivotValue>>();
        var filters = parameters["filters"]?.ToObject<List<string>>();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sourceSheet) || string.IsNullOrEmpty(sourceRange) ||
            string.IsNullOrEmpty(targetSheet) || string.IsNullOrEmpty(targetCell) || rows == null)
            throw new ArgumentException("Required parameters missing for pivot table");

        await _excelHandler.CreatePivotTableAsync(filepath, sourceSheet, sourceRange, targetSheet, targetCell, rows, columns, values, filters);
        return new JObject { ["success"] = true };
    }

    // Protection
    public static async Task<JObject> HandleProtectionAdd(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var password = parameters["password"]?.ToString();
        var allowFormatting = parameters["allow_formatting"]?.ToObject<bool>() ?? false;
        var allowSorting = parameters["allow_sorting"]?.ToObject<bool>() ?? false;

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("filepath and sheet_name required");

        await _excelHandler.AddProtectionAsync(filepath, sheetName, range, password, allowFormatting, allowSorting);
        return new JObject { ["success"] = true };
    }

    // Data validation
    public static async Task<JObject> HandleValidationAdd(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var range = parameters["range"]?.ToString();
        var validationType = parameters["validation_type"]?.ToString();
        var criteria = parameters["criteria"]?.ToObject<ValidationCriteria>();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(range) ||
            string.IsNullOrEmpty(validationType) || criteria == null)
            throw new ArgumentException("All parameters required for data validation");

        await _excelHandler.AddDataValidationAsync(filepath, sheetName, range, validationType, criteria);
        return new JObject { ["success"] = true };
    }

    // Comments
    public static async Task<JObject> HandleCommentAdd(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var cell = parameters["cell"]?.ToString();
        var text = parameters["text"]?.ToString();
        var author = parameters["author"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cell) || string.IsNullOrEmpty(text))
            throw new ArgumentException("filepath, sheet_name, cell, and text required");

        await _excelHandler.AddCommentAsync(filepath, sheetName, cell, text, author);
        return new JObject { ["success"] = true };
    }

    // Hyperlinks
    public static async Task<JObject> HandleHyperlinkAdd(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var cell = parameters["cell"]?.ToString();
        var url = parameters["url"]?.ToString();
        var displayText = parameters["display_text"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cell) || string.IsNullOrEmpty(url))
            throw new ArgumentException("filepath, sheet_name, cell, and url required");

        await _excelHandler.AddHyperlinkAsync(filepath, sheetName, cell, url, displayText);
        return new JObject { ["success"] = true };
    }

    // Images
    public static async Task<JObject> HandleImageAdd(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var imagePath = parameters["image_path"]?.ToString();
        var cell = parameters["cell"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(imagePath) || string.IsNullOrEmpty(cell))
            throw new ArgumentException("filepath, sheet_name, image_path, and cell required");

        await _excelHandler.AddImageAsync(filepath, sheetName, imagePath, cell);
        return new JObject { ["success"] = true };
    }

    // Find and replace
    public static async Task<JObject> HandleFindReplace(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var findText = parameters["find_text"]?.ToString();
        var replaceText = parameters["replace_text"]?.ToString();
        var range = parameters["range"]?.ToString();
        var matchCase = parameters["match_case"]?.ToObject<bool>() ?? false;
        var matchEntireCell = parameters["match_entire_cell"]?.ToObject<bool>() ?? false;

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(findText) || replaceText == null)
            throw new ArgumentException("filepath, sheet_name, find_text, and replace_text required");

        await _excelHandler.FindReplaceAsync(filepath, sheetName, findText, replaceText, range, matchCase, matchEntireCell);
        return new JObject { ["success"] = true };
    }

    // Add to McpServer's _toolHandlers dictionary
    public static void RegisterHandlers(Dictionary<string, Func<JObject, Task<JObject>>> handlers)
    {
        handlers["workbook_metadata"] = HandleWorkbookMetadata;
        handlers["io_import_csv"] = HandleDataImportCsv;
        handlers["io_export_csv"] = HandleDataExportCsv;
        handlers["format_range"] = HandleFormatRange;
        handlers["formula_apply"] = HandleFormulaApply;

        // Advanced formatting
        handlers["format_advanced"] = HandleFormatAdvanced;
        handlers["format_conditional"] = HandleFormatConditional;

        // Data manipulation
        handlers["data_sort"] = HandleDataSort;
        handlers["data_filter"] = HandleDataFilter;

        // Cell operations
        handlers["cell_merge"] = HandleCellMerge;
        handlers["cell_unmerge"] = HandleCellUnmerge;

        // Named ranges
        handlers["named_range_create"] = HandleNamedRangeCreate;

        // Charts and pivot tables
        handlers["chart_create"] = HandleChartCreate;
        handlers["pivot_create"] = HandlePivotCreate;

        // Protection and validation
        handlers["protection_add"] = HandleProtectionAdd;
        handlers["validation_add"] = HandleValidationAdd;

        // Comments and hyperlinks
        handlers["comment_add"] = HandleCommentAdd;
        handlers["hyperlink_add"] = HandleHyperlinkAdd;

        // Images and find/replace
        handlers["image_add"] = HandleImageAdd;
        handlers["find_replace"] = HandleFindReplace;
    }
}

// Extension methods for ExcelHandler
public static class ExcelHandlerExtensions
{
    public static async Task<WorkbookMetadata> GetWorkbookMetadataAsync(this ExcelHandler handler, string filepath)
    {
        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            return new WorkbookMetadata
            {
                FileName = Path.GetFileName(filepath),
                SheetCount = package.Workbook.Worksheets.Count,
                Sheets = package.Workbook.Worksheets.Select(ws => ws.Name).ToList()
            };
        });
    }

    public static async Task ImportCsvAsync(this ExcelHandler handler, string csvPath, string excelPath, string sheetName, bool hasHeader)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var worksheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName);

            using var reader = new StreamReader(csvPath);
            var csvData = reader.ReadToEnd();
            var lines = csvData.Split('\n');

            for (int row = 0; row < lines.Length; row++)
            {
                var cells = lines[row].Split(',');
                for (int col = 0; col < cells.Length; col++)
                {
                    worksheet.Cells[row + 1, col + 1].Value = cells[col].Trim('"');
                }
            }
            package.Save();
        });
    }

    public static async Task ExportCsvAsync(this ExcelHandler handler, string excelPath, string sheetName, string csvPath)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            using var writer = new StreamWriter(csvPath);
            for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
            {
                var line = string.Join(",", Enumerable.Range(worksheet.Dimension.Start.Column, worksheet.Dimension.End.Column)
                    .Select(col => $"\"{worksheet.Cells[row, col].Value?.ToString() ?? ""}\""));
                writer.WriteLine(line);
            }
        });
    }

    public static async Task FormatRangeAsync(this ExcelHandler handler, string filepath, string sheetName, string startCell, string endCell, bool bold, string? fillColor)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var range = worksheet.Cells[startCell + ":" + endCell];
            if (bold) range.Style.Font.Bold = true;
            if (!string.IsNullOrEmpty(fillColor))
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(fillColor));
            }
            package.Save();
        });
    }

    public static async Task ApplyFormulaAsync(this ExcelHandler handler, string filepath, string sheetName, string cell, string formula)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            worksheet.Cells[cell].Formula = formula;
            package.Save();
        });
    }
}

public class WorkbookMetadata
{
    public string? FileName { get; set; }
    public int SheetCount { get; set; }
    public List<string>? Sheets { get; set; }
}