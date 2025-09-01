using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Excel_mcp_dotnet;

public class McpServer
{
    private readonly ExcelHandler _excelHandler;
    private readonly Dictionary<string, Func<JObject, Task<JObject>>> _toolHandlers;
    private readonly ILogger<McpServer> _logger;

    public McpServer(ILogger<McpServer> logger)
    {
        _logger = logger;
        _excelHandler = new ExcelHandler();
        _toolHandlers = new Dictionary<string, Func<JObject, Task<JObject>>>
        {
            ["initialize"] = HandleInitialize,
            ["tools/list"] = HandleToolsList,
            ["tools/call"] = HandleToolCall,
        };
        
        // Register all the advanced tool handlers
        ToolHandlers.RegisterHandlers(_toolHandlers);
    }

    public async Task StartAsync(string[] args)
    {
        _logger.LogInformation("MCP Server started, listening on stdin/stdout");
        
        using var stdin = Console.OpenStandardInput();
        using var reader = new StreamReader(stdin, Encoding.UTF8);
        
        string? line;
        while ((line = await reader.ReadLineAsync()) != null)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            
            try
            {
                await HandleRequestAsync(line);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing request: {Request}", line);
                await WriteErrorResponseAsync(null, -32000, ex.Message);
            }
        }
    }

    private async Task HandleRequestAsync(string requestLine)
    {
        try
        {
            var request = JObject.Parse(requestLine);
            var method = request["method"]?.ToString();
            var id = request["id"];
            var parameters = request["params"] as JObject ?? new JObject();

            _logger.LogInformation("Received request: method={Method}, id={Id}", method, id);

            if (method == null)
            {
                await WriteErrorResponseAsync(id, -32600, "Invalid Request: missing method");
                return;
            }

            if (!_toolHandlers.ContainsKey(method))
            {
                await WriteErrorResponseAsync(id, -32601, $"Method not found: {method}");
                return;
            }

            var result = await _toolHandlers[method](parameters);
            await WriteSuccessResponseAsync(id, result);
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "JSON parsing error");
            await WriteErrorResponseAsync(null, -32700, "Parse error");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error handling request");
            await WriteErrorResponseAsync(null, -32000, ex.Message);
        }
    }

    private async Task WriteSuccessResponseAsync(JToken? id, JObject result)
    {
        var response = new JObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["result"] = result
        };
        
        await WriteResponseAsync(response);
    }

    private async Task WriteErrorResponseAsync(JToken? id, int code, string message)
    {
        var response = new JObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["error"] = new JObject
            {
                ["code"] = code,
                ["message"] = message
            }
        };
        
        await WriteResponseAsync(response);
    }

    private async Task WriteResponseAsync(JObject response)
    {
        var responseJson = JsonConvert.SerializeObject(response, Formatting.None);
        await Console.Out.WriteLineAsync(responseJson);
        await Console.Out.FlushAsync();
    }

    private Task<JObject> HandleInitialize(JObject parameters)
    {
        var result = new JObject
        {
            ["protocolVersion"] = "2024-11-05",
            ["capabilities"] = new JObject
            {
                ["tools"] = new JObject(),
                ["logging"] = new JObject()
            },
            ["serverInfo"] = new JObject
            {
                ["name"] = "excel-mcp-server",
                ["version"] = "1.0.0"
            }
        };

        _logger.LogInformation("MCP Server initialized with client");
        return Task.FromResult(result);
    }

    private Task<JObject> HandleToolsList(JObject parameters)
    {
        var tools = new JArray();

        // Define ALL available tools with proper schemas
        var toolDefinitions = new[]
        {
            // Core workbook operations
            new { name = "workbook-create", description = "Create a new Excel workbook", schema = CreateSimpleSchema("filepath") },
            new { name = "workbook-metadata", description = "Get workbook metadata", schema = CreateSimpleSchema("filepath") },
            
            // Worksheet operations
            new { name = "worksheet-create", description = "Create new worksheet", schema = CreateSchema(new[] { "filepath", "sheet_name" }) },
            new { name = "worksheet-delete", description = "Delete a worksheet", schema = CreateSchema(new[] { "filepath", "sheet_name" }) },
            new { name = "worksheet-rename", description = "Rename a worksheet", schema = CreateSchema(new[] { "filepath", "old_name", "new_name" }) },
            
            // Data operations
            new { name = "data-write", description = "Write 2D array data to worksheet", schema = CreateDataWriteSchema() },
            new { name = "data-read", description = "Read data from worksheet", schema = CreateDataReadSchema() },
            new { name = "cell-write", description = "Write value to a single cell", schema = CreateCellWriteSchema() },
            
            // Import/Export operations
            new { name = "io-import-csv", description = "Import CSV data to Excel", schema = CreateSchema(new[] { "csv_path", "excel_path" }, new[] { "sheet_name", "has_header" }) },
            new { name = "io-export-csv", description = "Export Excel data to CSV", schema = CreateSchema(new[] { "excel_path", "sheet_name", "csv_path" }) },
            
            // Formatting operations
            new { name = "format-range", description = "Apply basic formatting to a cell range", schema = CreateFormatRangeSchema() },
            new { name = "format-advanced", description = "Apply advanced formatting (fonts, borders, fills, alignment)", schema = CreateAdvancedFormatSchema() },
            new { name = "format-conditional", description = "Apply conditional formatting to a range", schema = CreateConditionalFormatSchema() },
            
            // Formula operations
            new { name = "formula-apply", description = "Apply a formula to a cell", schema = CreateSchema(new[] { "filepath", "sheet_name", "cell", "formula" }) },
            
            // Data manipulation
            new { name = "data-sort", description = "Sort data by one or multiple columns", schema = CreateSortSchema() },
            new { name = "data-filter", description = "Apply filters to a data range", schema = CreateFilterSchema() },
            new { name = "find-replace", description = "Find and replace text in worksheet", schema = CreateFindReplaceSchema() },
            
            // Cell operations
            new { name = "range-merge", description = "Merge cells in a range", schema = CreateSchema(new[] { "filepath", "sheet_name", "range" }) },
            new { name = "range-unmerge", description = "Unmerge cells in a range", schema = CreateSchema(new[] { "filepath", "sheet_name", "range" }) },
            
            // Advanced Excel features
            new { name = "table-create", description = "Create an Excel table with auto-filters", schema = CreateTableSchema() },
            new { name = "chart-create", description = "Create a chart in Excel", schema = CreateChartSchema() },
            new { name = "pivot-create", description = "Create a pivot table for data analysis", schema = CreatePivotSchema() },
            
            // Named ranges
            new { name = "named-range-create", description = "Create a named range for easy reference", schema = CreateSchema(new[] { "filepath", "name", "sheet_name", "range" }) },
            
            // Data validation
            new { name = "validation-add", description = "Add data validation to a range", schema = CreateValidationSchema() },
            
            // Protection
            new { name = "protection-add", description = "Add protection to worksheet or range", schema = CreateProtectionSchema() },
            
            // Comments and annotations
            new { name = "comment-add", description = "Add a comment to a cell", schema = CreateSchema(new[] { "filepath", "sheet_name", "cell", "text" }, new[] { "author" }) },
            new { name = "hyperlink-add", description = "Add a hyperlink to a cell", schema = CreateSchema(new[] { "filepath", "sheet_name", "cell", "url" }, new[] { "display_text" }) },
            
            // Images
            new { name = "image-add", description = "Add an image to a worksheet", schema = CreateSchema(new[] { "filepath", "sheet_name", "image_path", "cell" }) },
            
            // Row/Column operations
            new { name = "rows-insert", description = "Insert rows at specified position", schema = CreateRowColumnSchema() },
            new { name = "columns-insert", description = "Insert columns at specified position", schema = CreateRowColumnSchema() },
            new { name = "rows-delete", description = "Delete rows at specified position", schema = CreateRowColumnSchema() },
            new { name = "columns-delete", description = "Delete columns at specified position", schema = CreateRowColumnSchema() },
            
            // VBA operations
            new { name = "vba-read", description = "Read VBA code from workbook", schema = CreateSimpleSchema("filepath") },
            new { name = "vba-write", description = "Write VBA code to workbook", schema = CreateSchema(new[] { "filepath", "vba_code" }) },
            new { name = "vba-modules", description = "List VBA modules in workbook", schema = CreateSimpleSchema("filepath") },
            new { name = "vba-module-read", description = "Read specific VBA module", schema = CreateSchema(new[] { "filepath", "module_name" }) },
            new { name = "vba-module-write", description = "Write to specific VBA module", schema = CreateSchema(new[] { "filepath", "module_name", "vba_code" }) },
            new { name = "vba-module-delete", description = "Delete a VBA module", schema = CreateSchema(new[] { "filepath", "module_name" }) },
            
            // Server status
            new { name = "server-status", description = "Get MCP server status and information", schema = CreateSchema(new string[0]) }
        };

        foreach (var tool in toolDefinitions)
        {
            tools.Add(new JObject
            {
                ["name"] = tool.name,
                ["description"] = tool.description,
                ["inputSchema"] = tool.schema
            });
        }

        return Task.FromResult(new JObject { ["tools"] = tools });
    }

    // Helper methods to create schemas
    private JObject CreateSimpleSchema(string requiredParam)
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                [requiredParam] = new JObject { ["type"] = "string", ["description"] = $"The {requiredParam} parameter" }
            },
            ["required"] = new JArray(requiredParam)
        };
    }

    private JObject CreateSchema(string[] required, string[]? optional = null)
    {
        var properties = new JObject();
        foreach (var param in required)
        {
            properties[param] = new JObject { ["type"] = "string", ["description"] = $"The {param} parameter" };
        }
        if (optional != null)
        {
            foreach (var param in optional)
            {
                properties[param] = new JObject { ["type"] = "string", ["description"] = $"The {param} parameter (optional)" };
            }
        }

        return new JObject
        {
            ["type"] = "object",
            ["properties"] = properties,
            ["required"] = new JArray(required)
        };
    }

    private JObject CreateDataWriteSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["data"] = new JObject { ["type"] = "array", ["description"] = "2D array of data to write" },
                ["start_cell"] = new JObject { ["type"] = "string", ["description"] = "Starting cell (e.g., 'A1')", ["default"] = "A1" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "data")
        };
    }

    private JObject CreateDataReadSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["start_cell"] = new JObject { ["type"] = "string", ["description"] = "Starting cell (e.g., 'A1')", ["default"] = "A1" },
                ["end_cell"] = new JObject { ["type"] = "string", ["description"] = "Ending cell (optional)" }
            },
            ["required"] = new JArray("filepath", "sheet_name")
        };
    }

    private JObject CreateCellWriteSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["cell"] = new JObject { ["type"] = "string", ["description"] = "Cell address (e.g., 'A1')" },
                ["value"] = new JObject { ["description"] = "Value to write to the cell" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "cell", "value")
        };
    }

    private JObject CreateFormatRangeSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["start_cell"] = new JObject { ["type"] = "string", ["description"] = "Starting cell" },
                ["end_cell"] = new JObject { ["type"] = "string", ["description"] = "Ending cell" },
                ["bold"] = new JObject { ["type"] = "boolean", ["description"] = "Make text bold" },
                ["fill_color"] = new JObject { ["type"] = "string", ["description"] = "Fill color (hex format)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "start_cell", "end_cell")
        };
    }

    private JObject CreateAdvancedFormatSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Cell range to format" },
                ["formatting"] = new JObject { ["type"] = "object", ["description"] = "Advanced formatting options (font, border, fill, alignment)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range", "formatting")
        };
    }

    private JObject CreateConditionalFormatSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Cell range for conditional formatting" },
                ["rule_type"] = new JObject { ["type"] = "string", ["description"] = "Type of conditional formatting rule" },
                ["condition"] = new JObject { ["type"] = "object", ["description"] = "Condition for formatting" },
                ["format"] = new JObject { ["type"] = "object", ["description"] = "Format to apply when condition is met" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range", "rule_type", "condition", "format")
        };
    }

    private JObject CreateSortSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range to sort" },
                ["sort_by"] = new JObject { ["type"] = "array", ["description"] = "Array of sort columns with order" },
                ["ascending"] = new JObject { ["type"] = "boolean", ["description"] = "Sort ascending (default: true)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range", "sort_by")
        };
    }

    private JObject CreateFilterSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range to apply filters" },
                ["filters"] = new JObject { ["type"] = "object", ["description"] = "Filter criteria" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range")
        };
    }

    private JObject CreateFindReplaceSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["find_text"] = new JObject { ["type"] = "string", ["description"] = "Text to find" },
                ["replace_text"] = new JObject { ["type"] = "string", ["description"] = "Text to replace with" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range to search (optional)" },
                ["match_case"] = new JObject { ["type"] = "boolean", ["description"] = "Match case (default: false)" },
                ["match_entire_cell"] = new JObject { ["type"] = "boolean", ["description"] = "Match entire cell (default: false)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "find_text", "replace_text")
        };
    }

    private JObject CreateTableSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range for the table" },
                ["table_name"] = new JObject { ["type"] = "string", ["description"] = "Name for the table" },
                ["has_headers"] = new JObject { ["type"] = "boolean", ["description"] = "Table has headers (default: true)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range")
        };
    }

    private JObject CreateChartSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["data_range"] = new JObject { ["type"] = "string", ["description"] = "Data range for the chart" },
                ["chart_type"] = new JObject { ["type"] = "string", ["description"] = "Type of chart (Column, Line, Pie, etc.)" },
                ["target_cell"] = new JObject { ["type"] = "string", ["description"] = "Cell where to place the chart" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "data_range", "chart_type", "target_cell")
        };
    }

    private JObject CreatePivotSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["source_sheet"] = new JObject { ["type"] = "string", ["description"] = "Source worksheet name" },
                ["source_range"] = new JObject { ["type"] = "string", ["description"] = "Source data range" },
                ["target_sheet"] = new JObject { ["type"] = "string", ["description"] = "Target worksheet name" },
                ["target_cell"] = new JObject { ["type"] = "string", ["description"] = "Target cell for pivot table" },
                ["rows"] = new JObject { ["type"] = "array", ["description"] = "Row fields" },
                ["columns"] = new JObject { ["type"] = "array", ["description"] = "Column fields" },
                ["values"] = new JObject { ["type"] = "array", ["description"] = "Value fields" },
                ["filters"] = new JObject { ["type"] = "array", ["description"] = "Filter fields" }
            },
            ["required"] = new JArray("filepath", "source_sheet", "source_range", "target_sheet", "target_cell", "rows")
        };
    }

    private JObject CreateValidationSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range for validation" },
                ["validation_type"] = new JObject { ["type"] = "string", ["description"] = "Type of validation (list, number, date, etc.)" },
                ["criteria"] = new JObject { ["type"] = "object", ["description"] = "Validation criteria" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "range", "validation_type", "criteria")
        };
    }

    private JObject CreateProtectionSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["range"] = new JObject { ["type"] = "string", ["description"] = "Range to protect (optional)" },
                ["password"] = new JObject { ["type"] = "string", ["description"] = "Protection password (optional)" },
                ["allow_formatting"] = new JObject { ["type"] = "boolean", ["description"] = "Allow formatting (default: false)" },
                ["allow_sorting"] = new JObject { ["type"] = "boolean", ["description"] = "Allow sorting (default: false)" }
            },
            ["required"] = new JArray("filepath", "sheet_name")
        };
    }

    private JObject CreateRowColumnSchema()
    {
        return new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["filepath"] = new JObject { ["type"] = "string", ["description"] = "Path to the Excel workbook" },
                ["sheet_name"] = new JObject { ["type"] = "string", ["description"] = "Name of the worksheet" },
                ["position"] = new JObject { ["type"] = "integer", ["description"] = "Position to insert/delete (1-based)" },
                ["count"] = new JObject { ["type"] = "integer", ["description"] = "Number of rows/columns (default: 1)" }
            },
            ["required"] = new JArray("filepath", "sheet_name", "position")
        };
    }

    private async Task<JObject> HandleToolCall(JObject parameters)
    {
        var toolName = parameters["name"]?.ToString();
        var arguments = parameters["arguments"] as JObject ?? new JObject();

        if (string.IsNullOrEmpty(toolName))
        {
            throw new ArgumentException("Tool name is required");
        }

        // Convert tool name from kebab-case to snake_case for handler lookup
        var handlerName = toolName.Replace("-", "_");
        
        // Handle special cases for core tools not in ToolHandlers
        switch (toolName)
        {
            case "workbook-create":
                return await HandleWorkbookCreate(arguments);
            case "worksheet-create":
                return await HandleWorksheetCreate(arguments);
            case "worksheet-delete":
                return await HandleWorksheetDelete(arguments);
            case "worksheet-rename":
                return await HandleWorksheetRename(arguments);
            case "data-write":
                return await HandleDataWrite(arguments);
            case "data-read":
                return await HandleDataRead(arguments);
            case "cell-write":
                return await HandleCellWrite(arguments);
            case "server-status":
                return await HandleServerStatus(arguments);
        }

        // Try to find handler in registered tool handlers
        if (_toolHandlers.TryGetValue(handlerName, out var handler))
        {
            return await handler(arguments);
        }

        throw new ArgumentException($"Unknown tool: {toolName}");
    }

    // Tool handlers
    private async Task<JObject> HandleWorkbookCreate(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        if (string.IsNullOrEmpty(filepath)) throw new ArgumentException("filepath required");

        await _excelHandler.CreateWorkbookAsync(filepath);
        return new JObject { ["success"] = true };
    }

    private async Task<JObject> HandleWorksheetCreate(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName)) throw new ArgumentException("filepath and sheet_name required");

        await _excelHandler.CreateWorksheetAsync(filepath, sheetName);
        return new JObject { ["success"] = true };
    }

    private async Task<JObject> HandleWorksheetDelete(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName)) throw new ArgumentException("filepath and sheet_name required");

        await _excelHandler.DeleteWorksheetAsync(filepath, sheetName);
        return new JObject { ["success"] = true };
    }

    private async Task<JObject> HandleWorksheetRename(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var oldName = parameters["old_name"]?.ToString();
        var newName = parameters["new_name"]?.ToString();
        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName)) throw new ArgumentException("filepath, old_name, and new_name required");

        await _excelHandler.RenameWorksheetAsync(filepath, oldName, newName);
        return new JObject { ["success"] = true };
    }

    private async Task<JObject> HandleDataWrite(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var data = parameters["data"]?.ToObject<List<List<object>>>();
        var startCell = parameters["start_cell"]?.ToString() ?? "A1";

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || data == null) throw new ArgumentException("filepath, sheet_name, and data required");

        await _excelHandler.WriteDataAsync(filepath, sheetName, data, startCell);
        return new JObject { ["success"] = true };
    }

    private async Task<JObject> HandleDataRead(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var startCell = parameters["start_cell"]?.ToString() ?? "A1";
        var endCell = parameters["end_cell"]?.ToString();

        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName)) throw new ArgumentException("filepath and sheet_name required");

        var data = await _excelHandler.ReadDataAsync(filepath, sheetName, startCell, endCell);
        return new JObject { ["data"] = JArray.FromObject(data) };
    }

    private async Task<JObject> HandleCellWrite(JObject parameters)
    {
        var filepath = parameters["filepath"]?.ToString();
        var sheetName = parameters["sheet_name"]?.ToString();
        var cell = parameters["cell"]?.ToString();
        var value = parameters["value"];
        if (string.IsNullOrEmpty(filepath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cell)) throw new ArgumentException("filepath, sheet_name, and cell required");

        await _excelHandler.WriteCellAsync(filepath, sheetName, cell, value);
        return new JObject { ["success"] = true };
    }

    private Task<JObject> HandleServerStatus(JObject parameters)
    {
        return Task.FromResult(new JObject 
        { 
            ["status"] = "running",
            ["version"] = "1.0.0",
            ["description"] = "Excel MCP Server for .NET",
            ["capabilities"] = new JArray("workbook_operations", "worksheet_operations", "data_operations", "cell_operations")
        });
    }
}