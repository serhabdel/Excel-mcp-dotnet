using Excel_mcp_dotnet;
using Microsoft.Extensions.Logging;

// Configure logging to stderr to avoid interfering with MCP protocol on stdout
var loggerFactory = LoggerFactory.Create(builder => 
    builder.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace));
var logger = loggerFactory.CreateLogger<McpServer>();

var server = new McpServer(logger);
await server.StartAsync(args);
