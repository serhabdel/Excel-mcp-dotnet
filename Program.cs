using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Syncfusion.Licensing;

// Register your Syncfusion license key directly here
// Replace "YOUR_LICENSE_KEY" with your actual Syncfusion license key
SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");

var builder = Host.CreateApplicationBuilder(args);

// Configure logging to stderr for MCP protocol compatibility
builder.Logging.AddConsole(consoleLogOptions =>
{
    consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
});

// Add MCP server with stdio transport
builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
