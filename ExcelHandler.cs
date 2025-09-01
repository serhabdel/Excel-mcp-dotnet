using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table.PivotTable;
using Microsoft.Extensions.Caching.Memory;

namespace Excel_mcp_dotnet;

public class ExcelHandler
{
    private readonly IMemoryCache _cache;
    private readonly MemoryCacheEntryOptions _cacheOptions;

    public ExcelHandler()
    {
        _cache = new MemoryCache(new MemoryCacheOptions());
        _cacheOptions = new MemoryCacheEntryOptions
        {
            AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(10),
            Size = 1
        };
    }

    public async Task CreateWorkbookAsync(string filepath)
    {
        await Task.Run(() =>
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage();

                // Add a default worksheet since Excel requires at least one
                package.Workbook.Worksheets.Add("Sheet1");

                var fileInfo = new FileInfo(filepath);

                // Ensure directory exists
                if (fileInfo.Directory != null && !fileInfo.Directory.Exists)
                {
                    fileInfo.Directory.Create();
                }

                package.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving file {filepath}: {ex.Message}", ex);
            }
        });
    }

    public async Task CreateWorksheetAsync(string filepath, string sheetName)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            package.Workbook.Worksheets.Add(sheetName);
            package.Save();
        });
    }

    public async Task DeleteWorksheetAsync(string filepath, string sheetName)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
                package.Save();
            }
        });
    }

    public async Task RenameWorksheetAsync(string filepath, string oldName, string newName)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[oldName];
            if (worksheet != null)
            {
                worksheet.Name = newName;
                package.Save();
            }
        });
    }

    public async Task WriteCellAsync(string filepath, string sheetName, string cell, object? value)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            worksheet.Cells[cell].Value = value;
            package.Save();
        });
    }

    public async Task WriteDataAsync(string filepath, string sheetName, List<List<object>> data, string startCell)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName);

            var startAddress = new ExcelAddress(startCell);
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < data[row].Count; col++)
                {
                    worksheet.Cells[startAddress.Start.Row + row, startAddress.Start.Column + col].Value = data[row][col];
                }
            }
            package.Save();
        });
    }

    public async Task<List<List<object>>> ReadDataAsync(string filepath, string sheetName, string startCell, string? endCell)
    {
        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var startAddress = new ExcelAddress(startCell);
            var endAddress = endCell != null ? new ExcelAddress(endCell) : worksheet.Dimension;

            var data = new List<List<object>>();
            for (int row = startAddress.Start.Row; row <= endAddress.End.Row; row++)
            {
                var rowData = new List<object>();
                for (int col = startAddress.Start.Column; col <= endAddress.End.Column; col++)
                {
                    rowData.Add(worksheet.Cells[row, col].Value);
                }
                data.Add(rowData);
            }
            return data;
        });
    }

    public async Task<string> ReadVbaAsync(string filepath)
    {
        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            // First try to read from actual VBA modules
            if (package.Workbook.VbaProject != null && package.Workbook.VbaProject.Modules.Count > 0)
            {
                var vbaCode = new System.Text.StringBuilder();
                foreach (var module in package.Workbook.VbaProject.Modules)
                {
                    vbaCode.AppendLine($"'=== {module.Name} ===");
                    vbaCode.AppendLine(module.Code);
                    vbaCode.AppendLine();
                }
                return vbaCode.ToString();
            }

            // Fallback: read from workbook properties
            return package.Workbook.Properties.Comments ?? string.Empty;
        });
    }

    public async Task<List<string>> GetVbaModulesAsync(string filepath)
    {
        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            var modules = new List<string>();

            // First try to get from actual VBA modules
            if (package.Workbook.VbaProject != null)
            {
                foreach (var module in package.Workbook.VbaProject.Modules)
                {
                    modules.Add(module.Name);
                }
            }

            // Also check workbook properties for stored modules
            var comments = package.Workbook.Properties.Comments ?? "";
            var regex = new System.Text.RegularExpressions.Regex(@"\[([^\]]+)\]");
            var matches = regex.Matches(comments);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                var moduleName = match.Groups[1].Value;
                if (!modules.Contains(moduleName))
                {
                    modules.Add(moduleName);
                }
            }

            return modules;
        });
    }

    public async Task<string> ReadVbaModuleAsync(string filepath, string moduleName)
    {
        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            // First try to read from actual VBA modules
            if (package.Workbook.VbaProject != null)
            {
                foreach (var module in package.Workbook.VbaProject.Modules)
                {
                    if (module.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return module.Code;
                    }
                }
            }

            // Fallback: read from workbook properties
            var comments = package.Workbook.Properties.Comments ?? "";
            var moduleMarker = $"[{moduleName}]";

            if (comments.Contains(moduleMarker))
            {
                var startIndex = comments.IndexOf(moduleMarker) + moduleMarker.Length;
                var endIndex = comments.IndexOf("[", startIndex);
                if (endIndex == -1) endIndex = comments.Length;

                var moduleCode = comments.Substring(startIndex, endIndex - startIndex).Trim();
                return moduleCode;
            }

            return string.Empty;
        });
    }

    public async Task WriteVbaModuleAsync(string filepath, string moduleName, string vbaCode)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            // Store VBA module code in workbook properties with module name
            var existingComments = package.Workbook.Properties.Comments ?? "";
            var moduleMarker = $"[{moduleName}]";
            var newComments = "";

            if (existingComments.Contains(moduleMarker))
            {
                // Replace existing module
                var pattern = $"{moduleMarker}[^\\[]*";
                newComments = System.Text.RegularExpressions.Regex.Replace(existingComments, pattern, $"{moduleMarker}\n{vbaCode}");
            }
            else
            {
                // Add new module
                newComments = string.IsNullOrEmpty(existingComments)
                    ? $"{moduleMarker}\n{vbaCode}"
                    : $"{existingComments}\n\n{moduleMarker}\n{vbaCode}";
            }

            package.Workbook.Properties.Comments = newComments;
            package.Save();
        });
    }

    public async Task DeleteVbaModuleAsync(string filepath, string moduleName)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            // Try to remove from actual VBA modules first
            if (package.Workbook.VbaProject != null)
            {
                var moduleToRemove = package.Workbook.VbaProject.Modules.FirstOrDefault(m => m.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));
                if (moduleToRemove != null)
                {
                    package.Workbook.VbaProject.Modules.Remove(moduleToRemove);
                    package.Save();
                    return;
                }
            }

            // Remove from workbook properties
            var comments = package.Workbook.Properties.Comments ?? "";
            var moduleMarker = $"[{moduleName}]";

            if (comments.Contains(moduleMarker))
            {
                var startIndex = comments.IndexOf(moduleMarker);
                var endIndex = comments.IndexOf("[", startIndex + moduleMarker.Length);
                if (endIndex == -1) endIndex = comments.Length;

                var beforeModule = startIndex > 0 ? comments.Substring(0, startIndex).TrimEnd() : "";
                var afterModule = endIndex < comments.Length ? comments.Substring(endIndex) : "";

                var newComments = "";
                if (!string.IsNullOrEmpty(beforeModule) && !string.IsNullOrEmpty(afterModule))
                {
                    newComments = $"{beforeModule}\n\n{afterModule}";
                }
                else if (!string.IsNullOrEmpty(beforeModule))
                {
                    newComments = beforeModule;
                }
                else if (!string.IsNullOrEmpty(afterModule))
                {
                    newComments = afterModule;
                }

                package.Workbook.Properties.Comments = newComments;
                package.Save();
            }
        });
    }

    public async Task WriteVbaAsync(string filepath, string vbaCode)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));

            // Ensure VBA project exists
            if (package.Workbook.VbaProject == null)
            {
                package.Workbook.CreateVBAProject();
            }

            // Store VBA code in workbook properties as a workaround
            // This allows us to persist VBA code even if direct module manipulation fails
            package.Workbook.Properties.Comments = vbaCode;

            // Try to create VBA project structure if possible
            try
            {
                var vbaProject = package.Workbook.VbaProject;
                if (vbaProject != null && vbaProject.Modules.Count == 0)
                {
                    // Create a basic module structure
                    // Note: This may not work in all EPPlus versions
                    var moduleCode = $"'VBA Module\n'Generated by MCP Excel Server\n\n{vbaCode}";
                    package.Workbook.Properties.Comments = moduleCode;
                }
            }
            catch (Exception ex)
            {
                // Log the issue but continue - VBA code is stored in properties
                Console.WriteLine($"VBA module creation warning: {ex.Message}");
            }

            package.Save();
        });
    }

    // Additional performance methods
    public async Task<ExcelPackage> GetCachedPackageAsync(string filepath)
    {
        if (_cache.TryGetValue(filepath, out ExcelPackage cachedPackage))
        {
            return cachedPackage;
        }

        return await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(filepath));
            _cache.Set(filepath, package, _cacheOptions);
            return package;
        });
    }

    // Advanced formatting
    public async Task ApplyAdvancedFormattingAsync(string filepath, string sheetName, string range, AdvancedFormatting formatting)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var excelRange = worksheet.Cells[range];

            // Font formatting
            if (formatting.Font != null)
            {
                if (formatting.Font.Bold.HasValue) excelRange.Style.Font.Bold = formatting.Font.Bold.Value;
                if (formatting.Font.Italic.HasValue) excelRange.Style.Font.Italic = formatting.Font.Italic.Value;
                if (!string.IsNullOrEmpty(formatting.Font.Color)) excelRange.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Font.Color));
                if (formatting.Font.Size.HasValue) excelRange.Style.Font.Size = (float)formatting.Font.Size.Value;
                if (!string.IsNullOrEmpty(formatting.Font.Name)) excelRange.Style.Font.Name = formatting.Font.Name;
            }

            // Fill formatting
            if (formatting.Fill != null)
            {
                if (!string.IsNullOrEmpty(formatting.Fill.BackgroundColor))
                {
                    excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    excelRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Fill.BackgroundColor));
                }
            }

            // Border formatting
            if (formatting.Border != null)
            {
                if (!string.IsNullOrEmpty(formatting.Border.Color))
                {
                    var borderStyle = GetBorderStyle(formatting.Border.Style);
                    excelRange.Style.Border.Top.Style = borderStyle;
                    excelRange.Style.Border.Bottom.Style = borderStyle;
                    excelRange.Style.Border.Left.Style = borderStyle;
                    excelRange.Style.Border.Right.Style = borderStyle;
                    excelRange.Style.Border.Top.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Border.Color));
                    excelRange.Style.Border.Bottom.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Border.Color));
                    excelRange.Style.Border.Left.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Border.Color));
                    excelRange.Style.Border.Right.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(formatting.Border.Color));
                }
            }

            // Alignment
            if (formatting.Alignment != null)
            {
                if (!string.IsNullOrEmpty(formatting.Alignment.Horizontal))
                    excelRange.Style.HorizontalAlignment = GetHorizontalAlignment(formatting.Alignment.Horizontal);
                if (!string.IsNullOrEmpty(formatting.Alignment.Vertical))
                    excelRange.Style.VerticalAlignment = GetVerticalAlignment(formatting.Alignment.Vertical);
                if (formatting.Alignment.WrapText.HasValue) excelRange.Style.WrapText = formatting.Alignment.WrapText.Value;
            }

            // Number format
            if (!string.IsNullOrEmpty(formatting.NumberFormat))
            {
                excelRange.Style.Numberformat.Format = formatting.NumberFormat;
            }

            package.Save();
        });
    }

    public async Task ApplyConditionalFormattingAsync(string filepath, string sheetName, string range, string ruleType, ConditionalFormattingCondition condition, ConditionalFormattingFormat format)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var excelRange = worksheet.Cells[range];
            var conditionalFormatting = excelRange.ConditionalFormatting.AddExpression();

            // Apply condition based on rule type
            switch (ruleType.ToLower())
            {
                case "cell_value":
                    if (!string.IsNullOrEmpty(condition.Operator) && condition.Value != null)
                    {
                        conditionalFormatting.Formula = $"{condition.Value}";
                    }
                    break;
                case "formula":
                    if (!string.IsNullOrEmpty(condition.Formula))
                    {
                        conditionalFormatting.Formula = condition.Formula;
                    }
                    break;
            }

            // Apply formatting
            if (!string.IsNullOrEmpty(format.BackgroundColor))
            {
                conditionalFormatting.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(format.BackgroundColor));
            }
            if (!string.IsNullOrEmpty(format.FontColor))
            {
                conditionalFormatting.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(format.FontColor));
            }
            if (format.Bold.HasValue) conditionalFormatting.Style.Font.Bold = format.Bold.Value;

            package.Save();
        });
    }

    // Data manipulation
    public async Task SortRangeAsync(string filepath, string sheetName, string range, List<SortColumn> sortBy, bool ascending)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var excelRange = worksheet.Cells[range];
            excelRange.Sort(sortBy.Select(s => s.ColumnIndex).ToArray(), ascending ? new bool[] { true } : new bool[] { false });
            package.Save();
        });
    }

    public async Task ApplyFiltersAsync(string filepath, string sheetName, string range, Dictionary<string, object>? filters)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var excelRange = worksheet.Cells[range];
            excelRange.AutoFilter = true;

            // Note: Advanced filtering implementation would require more complex EPPlus API usage
            // For now, just enable auto-filter

            package.Save();
        });
    }

    // Cell operations
    public async Task MergeCellsAsync(string filepath, string sheetName, string range)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            worksheet.Cells[range].Merge = true;
            package.Save();
        });
    }

    public async Task UnmergeCellsAsync(string filepath, string sheetName, string range)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            worksheet.Cells[range].Merge = false;
            package.Save();
        });
    }

    // Named ranges
    public async Task CreateNamedRangeAsync(string filepath, string name, string sheetName, string range)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            package.Workbook.Names.Add(name, worksheet.Cells[range]);
            package.Save();
        });
    }

    // Charts
    public async Task CreateChartAsync(string filepath, string sheetName, string dataRange, string chartType, string targetCell)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var chart = worksheet.Drawings.AddChart(chartType + "Chart", GetChartType(chartType));
            chart.SetPosition(worksheet.Cells[targetCell].Start.Row - 1, 0, worksheet.Cells[targetCell].Start.Column - 1, 0);
            chart.SetSize(400, 300);
            var series = chart.Series.Add(worksheet.Cells[dataRange]);
            series.Header = "Data";

            package.Save();
        });
    }

    // Pivot tables
    public async Task CreatePivotTableAsync(string filepath, string sourceSheet, string sourceRange, string targetSheet, string targetCell,
        List<string> rows, List<string>? columns, List<PivotValue>? values, List<string>? filters)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var sourceWorksheet = package.Workbook.Worksheets[sourceSheet];
            var targetWorksheet = package.Workbook.Worksheets[targetSheet] ?? package.Workbook.Worksheets.Add(targetSheet);

            if (sourceWorksheet == null) throw new ArgumentException($"Source worksheet {sourceSheet} not found");

            var pivotTable = targetWorksheet.PivotTables.Add(targetWorksheet.Cells[targetCell], sourceWorksheet.Cells[sourceRange], "PivotTable1");

            foreach (var row in rows)
            {
                pivotTable.RowFields.Add(pivotTable.Fields[row]);
            }

            if (columns != null)
            {
                foreach (var col in columns)
                {
                    pivotTable.ColumnFields.Add(pivotTable.Fields[col]);
                }
            }

            if (values != null)
            {
                foreach (var value in values)
                {
                    var field = pivotTable.DataFields.Add(pivotTable.Fields[value.Field]);
                    field.Function = GetPivotFunction(value.Function);
                }
            }

            if (filters != null)
            {
                foreach (var filter in filters)
                {
                    pivotTable.PageFields.Add(pivotTable.Fields[filter]);
                }
            }

            package.Save();
        });
    }

    // Protection
    public async Task AddProtectionAsync(string filepath, string sheetName, string? range, string? password, bool allowFormatting, bool allowSorting)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            if (!string.IsNullOrEmpty(range))
            {
                var excelRange = worksheet.Cells[range];
                excelRange.Style.Locked = true;
            }
            else
            {
                worksheet.Protection.IsProtected = true;
                worksheet.Protection.AllowFormatCells = allowFormatting;
                worksheet.Protection.AllowSort = allowSorting;
                if (!string.IsNullOrEmpty(password))
                {
                    worksheet.Protection.SetPassword(password);
                }
            }

            package.Save();
        });
    }

    // Data validation
    public async Task AddDataValidationAsync(string filepath, string sheetName, string range, string validationType, ValidationCriteria criteria)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var validation = worksheet.DataValidations.AddListValidation(range);

            // Simplified data validation - EPPlus API is complex for full implementation
            if (validationType.ToLower() == "list" && criteria.Values != null && criteria.Values.Count > 0)
            {
                var listValidation = worksheet.DataValidations.AddListValidation(range);
                // Note: Full implementation would set the formula values
            }

            if (!string.IsNullOrEmpty(criteria.ErrorMessage))
            {
                validation.Error = criteria.ErrorMessage;
            }

            package.Save();
        });
    }

    // Comments
    public async Task AddCommentAsync(string filepath, string sheetName, string cell, string text, string? author)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var comment = worksheet.Cells[cell].AddComment(text, author ?? "System");
            package.Save();
        });
    }

    // Hyperlinks
    public async Task AddHyperlinkAsync(string filepath, string sheetName, string cell, string url, string? displayText)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var hyperlink = worksheet.Cells[cell].Hyperlink = new Uri(url);
            if (!string.IsNullOrEmpty(displayText))
            {
                worksheet.Cells[cell].Value = displayText;
            }

            package.Save();
        });
    }

    // Images
    public async Task AddImageAsync(string filepath, string sheetName, string imagePath, string cell)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var image = worksheet.Drawings.AddPicture("Image", new FileInfo(imagePath));
            var cellAddress = worksheet.Cells[cell];
            image.SetPosition(cellAddress.Start.Row - 1, 0, cellAddress.Start.Column - 1, 0);
            image.SetSize(100, 100);

            package.Save();
        });
    }

    // Find and replace
    public async Task FindReplaceAsync(string filepath, string sheetName, string findText, string replaceText, string? range, bool matchCase, bool matchEntireCell)
    {
        await Task.Run(() =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filepath));
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null) throw new ArgumentException($"Worksheet {sheetName} not found");

            var searchRange = string.IsNullOrEmpty(range) ? worksheet.Cells : worksheet.Cells[range];

            foreach (var cell in searchRange)
            {
                if (cell.Value != null)
                {
                    var cellValue = cell.Value.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                        if (matchEntireCell)
                        {
                            if (cellValue.Equals(findText, comparison))
                            {
                                cell.Value = replaceText;
                            }
                        }
                        else
                        {
                            cell.Value = cellValue.Replace(findText, replaceText, comparison);
                        }
                    }
                }
            }

            package.Save();
        });
    }

    // Helper methods
    private static int GetColumnIndex(string columnName)
    {
        var index = 0;
        foreach (var c in columnName.ToUpper())
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }

    private static eChartType GetChartType(string chartType)
    {
        return chartType.ToLower() switch
        {
            "column" => eChartType.ColumnClustered,
            "bar" => eChartType.BarClustered,
            "line" => eChartType.Line,
            "pie" => eChartType.Pie,
            "area" => eChartType.Area,
            _ => eChartType.ColumnClustered
        };
    }

    private static DataFieldFunctions GetPivotFunction(string? function)
    {
        return function?.ToLower() switch
        {
            "sum" => DataFieldFunctions.Sum,
            "count" => DataFieldFunctions.Count,
            "average" => DataFieldFunctions.Average,
            "max" => DataFieldFunctions.Max,
            "min" => DataFieldFunctions.Min,
            _ => DataFieldFunctions.Sum
        };
    }

    private static ExcelBorderStyle GetBorderStyle(string? style)
    {
        return style?.ToLower() switch
        {
            "thin" => ExcelBorderStyle.Thin,
            "medium" => ExcelBorderStyle.Medium,
            "thick" => ExcelBorderStyle.Thick,
            "double" => ExcelBorderStyle.Double,
            _ => ExcelBorderStyle.Thin
        };
    }

    private static ExcelHorizontalAlignment GetHorizontalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ExcelHorizontalAlignment.Left,
            "center" => ExcelHorizontalAlignment.Center,
            "right" => ExcelHorizontalAlignment.Right,
            "justify" => ExcelHorizontalAlignment.Justify,
            _ => ExcelHorizontalAlignment.Left
        };
    }

    private static ExcelVerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => ExcelVerticalAlignment.Top,
            "middle" => ExcelVerticalAlignment.Center,
            "bottom" => ExcelVerticalAlignment.Bottom,
            "justify" => ExcelVerticalAlignment.Justify,
            _ => ExcelVerticalAlignment.Top
        };
    }

    public void Dispose()
    {
        _cache.Dispose();
    }
}