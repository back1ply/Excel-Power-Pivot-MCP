using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;
using ExcelPowerPivotMcp.Common;
using ModelContextProtocol.Protocol;

namespace ExcelPowerPivotMcp;

sealed class Program
{
    static async Task Main(string[] args)
    {
        // Load configuration from environment variables
        var config = McpConfiguration.Load();
        
        // Log startup to stderr (visible to user, not sent to MCP client)
        Console.Error.WriteLine("[Excel Power Pivot MCP] Starting server...");
        
        var builder = Host.CreateApplicationBuilder(args);
        
        // CRITICAL: MCP requires stdout for JSON-RPC only!
        // Redirect all logging to stderr so errors are still visible but don't corrupt the protocol.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(options =>
        {
            options.LogToStandardErrorThreshold = LogLevel.Trace; // ALL logs go to stderr
        });

        
        // Register configuration
        builder.Services.AddSingleton(config);
        
        // Register PowerPivotConnectionManager singleton as interface
        // Uses a manager wrapper to avoid DI capturing a stale reference after Reset()
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IPowerPivotConnectionManager, ExcelPowerPivotMcp.Infrastructure.Services.PowerPivotConnectionManager>();

        // Register PowerPivotConnection as a factory that always returns the current instance
        // Register both the concrete type (for tools that need discovery methods) and interface (for infrastructure services)
        builder.Services.AddSingleton<PowerPivotConnection>(sp =>
            sp.GetRequiredService<ExcelPowerPivotMcp.Infrastructure.Services.IPowerPivotConnectionManager>().Current);

        builder.Services.AddSingleton<IPowerPivotConnection>(sp =>
            sp.GetRequiredService<ExcelPowerPivotMcp.Infrastructure.Services.IPowerPivotConnectionManager>().Current);
        
        // Register In-Memory Logger for MCP resource exposure
        var logProvider = new ExcelPowerPivotMcp.Infrastructure.Services.InMemoryLoggerProvider();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IInMemoryLogReader>(logProvider);
        builder.Logging.AddProvider(logProvider);

        // Infrastructure Services
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.ExcelStaService>();
        builder.Services.AddSingleton<Microsoft.Extensions.Hosting.IHostedService>(sp => sp.GetRequiredService<ExcelPowerPivotMcp.Infrastructure.Services.ExcelStaService>());
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IExcelDispatcher>(sp => sp.GetRequiredService<ExcelPowerPivotMcp.Infrastructure.Services.ExcelStaService>());

        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IExcelInteropService, ExcelPowerPivotMcp.Infrastructure.Services.ExcelInteropService>();

        // New decomposed COM services (SRP-compliant)
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IMeasureComService, ExcelPowerPivotMcp.Infrastructure.Services.MeasureComService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IRelationshipComService, ExcelPowerPivotMcp.Infrastructure.Services.RelationshipComService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.ITableComService, ExcelPowerPivotMcp.Infrastructure.Services.TableComService>();

        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IDmvService, ExcelPowerPivotMcp.Infrastructure.Services.AdoDmvService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Infrastructure.Services.IDaxService, ExcelPowerPivotMcp.Infrastructure.Services.AdoDaxService>();

        // Core Services
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.IMeasureService, ExcelPowerPivotMcp.Core.Services.MeasureService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.IRelationshipService, ExcelPowerPivotMcp.Core.Services.RelationshipService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.ITableService, ExcelPowerPivotMcp.Core.Services.TableService>();
        // NOTE: ICalculatedColumnService removed - Excel COM API doesn't support creating calculated columns
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.IModelMetadataService, ExcelPowerPivotMcp.Core.Services.ModelMetadataService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.IDataProfileService, ExcelPowerPivotMcp.Core.Services.DataProfileService>();
        builder.Services.AddSingleton<ExcelPowerPivotMcp.Core.Services.IDaxFormatterService, ExcelPowerPivotMcp.Core.Services.DaxFormatterService>();
        
        // Configure MCP server with stdio transport and auto-discover tools from this assembly
        builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "excel-powerpivot-mcp",
                    Version = typeof(Program).Assembly.GetName().Version?.ToString(3) ?? "1.0.0"
                };
            })
            .WithStdioServerTransport()
            .WithToolsFromAssembly()
            .WithResourcesFromAssembly()
            .WithPromptsFromAssembly()
            .AddCallToolFilter(next => async (context, cancellationToken) =>
            {
                try
                {
                    return await next(context, cancellationToken);
                }
                catch (Exception ex)
                {
                    // Log the full error to stderr for debugging
                    Console.Error.WriteLine($"[Error] Tool execution failed: {ex}");
                    
                    return new CallToolResult
                    {
                        Content = new[] { new TextContentBlock { Text = $"Error: {ex.Message}" } },
                        IsError = true
                    };
                }
            })
            .AddListToolsFilter(next => async (context, cancellationToken) =>
            {
                try
                {
                    return await next(context, cancellationToken);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[Error] List tools failed: {ex}");
                    throw; // Re-throw for list tools as we can't easily return a partial result
                }
            });
        
        var app = builder.Build();

        // Logger is now initialized automatically via PowerPivotConnectionManager constructor
        // This ensures proper DI pattern instead of manual service locator

        await app.RunAsync();
    }
}
