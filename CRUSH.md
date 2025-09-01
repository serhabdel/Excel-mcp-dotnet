# Excel MCP Server (.NET) - Development Guide

## Build & Test Commands

### Development
```bash
# Build project
dotnet build

# Run in development mode
dotnet run

# Run with specific test
echo '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"server-status","arguments":{}}}' | dotnet run
```

### Production Build
```bash
# Create optimized single-file executable
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true

# Test production build
./bin/Release/net8.0/linux-x64/publish/Excel-mcp-dotnet
```

### Testing
```bash
# Run MCP protocol tests
./test_mcp_flow.sh

# Manual tool testing
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | dotnet run
```

## Code Style Guidelines

### Imports & Using Statements
- Use `ImplicitUsings: enable` in project file
- Organize imports: System → Third-party → Project namespaces
- Avoid `using` statements for common .NET namespaces when implicit usings are enabled

### Naming Conventions
- **Classes**: PascalCase (e.g., `ExcelHandler`, `McpServer`)
- **Methods**: PascalCase (e.g., `CreateWorkbookAsync`, `HandleToolCall`)
- **Properties**: PascalCase (e.g., `FilePath`, `SheetName`)
- **Parameters**: camelCase (e.g., `filepath`, `sheetName`)
- **Private fields**: _camelCase (e.g., `_excelHandler`, `_logger`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `MAX_CACHE_SIZE`)

### Async/Await Patterns
- Use `Task.Run()` for CPU-bound Excel operations
- All public methods should be async and return `Task<T>`
- Use `await` consistently - avoid `.Result` or `.Wait()`
- Method names end with `Async` for async operations

### Error Handling
- Use try-catch blocks for file operations
- Wrap exceptions with context: `throw new Exception($"Error saving file {filepath}: {ex.Message}", ex)`
- Validate required parameters: `if (string.IsNullOrEmpty(filepath)) throw new ArgumentException("filepath required")`
- Log errors with structured logging: `_logger.LogError(ex, "Error processing request")`

### JSON & MCP Protocol
- Use `Newtonsoft.Json` for JSON serialization
- All tool responses must include `content` field with `JArray` of text objects
- Follow MCP protocol: `jsonrpc: "2.0"`, include `id`, proper error codes
- Use `JObject` for dynamic JSON handling

### Excel Operations (EPPlus)
- Set license context: `ExcelPackage.LicenseContext = LicenseContext.NonCommercial`
- Use `using` statements for `ExcelPackage` disposal
- Validate worksheet existence before operations
- Handle file paths with `FileInfo` and ensure directories exist

### Memory Management
- Use `IMemoryCache` for workbook caching (10-minute expiration)
- Implement `IDisposable` for cleanup
- Cache key format: `{filepath}|{operation}`

### Logging
- Use `Microsoft.Extensions.Logging`
- Log to stderr to avoid interfering with MCP protocol on stdout
- Include structured data: `LogInformation("Received request: method={Method}, id={Id}", method, id)`

### Tool Handler Pattern
- Register handlers in `ToolHandlers.RegisterHandlers()`
- Convert kebab-case tool names to snake_case for method lookup
- Return standardized response objects with success/data fields
- Use parameter validation and descriptive error messages