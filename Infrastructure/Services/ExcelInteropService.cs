using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.PowerPivot;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of Excel interop operations.
/// All COM operations are dispatched to the STA thread via IExcelDispatcher.
/// </summary>
public class ExcelInteropService : IExcelInteropService
{
    // NOTE: This service uses the concrete PowerPivotConnection class because it needs
    // access to workbook discovery methods (DiscoverWorkbooks, Connect, RefreshModel)
    // which are not part of the core IPowerPivotConnection interface used by COM services.
    private readonly PowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;
    private readonly McpConfiguration _config;
    private readonly ILogger<ExcelInteropService> _logger;

    public ExcelInteropService(
        PowerPivotConnection connection,
        IExcelDispatcher dispatcher,
        McpConfiguration config,
        ILogger<ExcelInteropService> logger)
    {
        _connection = connection;
        _dispatcher = dispatcher;
        _config = config;
        _logger = logger;
    }

    public async Task<List<PowerPivotWorkbook>> DiscoverWorkbooksAsync(CancellationToken ct = default)
    {
        var timeout = TimeSpan.FromMilliseconds(_config.ConnectionTimeoutMs);
        return await _dispatcher.RunAsync(() => _connection.DiscoverWorkbooks(), timeout, ct);
    }

    public async Task ConnectAsync(string workbookPath, CancellationToken ct = default)
    {
        int maxRetries = _config.ConnectionRetryCount;
        int baseDelayMs = _config.ConnectionRetryDelayMs;
        int maxDelayMs = _config.ConnectionMaxRetryDelayMs;
        var timeout = TimeSpan.FromMilliseconds(_config.ConnectionTimeoutMs);
        
        Exception? lastException = null;
        
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            try
            {
                await _dispatcher.RunAsync(() => _connection.Connect(workbookPath), timeout, ct);
                return; // Success
            }
            catch (Exception ex) when (attempt < maxRetries && !(ex is InvalidOperationException ioe && 
                (ioe.Message.Contains("Excel is not running") || 
                 ioe.Message.Contains("not found in Excel") ||
                 ioe.Message.Contains("does not have a Power Pivot"))))
            {
                // Retry on transient errors
                lastException = ex;
                int delay = Math.Min(baseDelayMs * (int)Math.Pow(2, attempt), maxDelayMs);
                _logger.LogWarning("Connection attempt {Attempt} failed: {Message}. Retrying in {DelayMs}ms...", 
                    attempt + 1, ex.Message, delay);
                
                await Task.Delay(delay, ct);
            }
        }
        
        throw lastException ?? new InvalidOperationException("Connection failed after retries.");
    }

    public async Task SaveWorkbookAsync(CancellationToken ct = default)
    {
        // Use longer timeout for save operations (may trigger dialogs)
        var timeout = TimeSpan.FromSeconds(30);
        await _dispatcher.RunAsync(() => _connection.SaveWorkbook(), timeout, ct);
    }

    public async Task RefreshModelAsync(CancellationToken ct = default)
    {
        // Refresh can be very long-running
        var timeout = TimeSpan.FromSeconds(_config.QueryTimeoutSeconds);
        await _dispatcher.RunAsync(() => _connection.RefreshModel(), timeout, ct);
    }

    public string? GetConnectedWorkbookPath()
    {
        return _connection.ConnectedWorkbook;
    }

    public bool IsConnected => _connection.IsConnected;

    public async Task<List<Dictionary<string, object?>>> GetExcelTablesAsync(CancellationToken ct = default)
    {
        var timeout = TimeSpan.FromSeconds(_config.DmvTimeoutSeconds);
        return await _dispatcher.RunAsync(() =>
        {
            if (!_connection.IsConnected || _connection.Workbook == null)
            {
                throw new InvalidOperationException("Must be connected to list tables.");
            }
            return ExcelComHelpers.GetExcelTables(_connection.Workbook, _logger);
        }, timeout, ct);
    }
}

