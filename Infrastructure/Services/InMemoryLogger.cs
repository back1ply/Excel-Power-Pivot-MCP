using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Keeps the last N log entries in memory for exposure via MCP resources.
/// Allows the Agent to "see what happened" when debugging issues.
/// </summary>
public interface IInMemoryLogReader
{
    string GetLogs();
}

public class InMemoryLoggerProvider : ILoggerProvider, IInMemoryLogReader
{
    private readonly ConcurrentQueue<string> _logQueue = new();
    private const int MaxLogCount = 100;

    public ILogger CreateLogger(string categoryName)
    {
        return new InMemoryLogger(categoryName, this);
    }

    internal void AddLog(string message)
    {
        _logQueue.Enqueue(message);
        
        // Trim old logs
        while (_logQueue.Count > MaxLogCount)
        {
            _logQueue.TryDequeue(out _);
        }
    }

    public string GetLogs()
    {
        var sb = new StringBuilder();
        sb.AppendLine(System.Globalization.CultureInfo.InvariantCulture, $"--- Server Logs (Last {_logQueue.Count} entries) ---");
        
        foreach (var log in _logQueue)
        {
            sb.AppendLine(log);
        }
        
        return sb.ToString();
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
    }
}

public class InMemoryLogger : ILogger
{
    private readonly string _categoryName;
    private readonly InMemoryLoggerProvider _provider;

    public InMemoryLogger(string categoryName, InMemoryLoggerProvider provider)
    {
        _categoryName = categoryName;
        _provider = provider;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => default!;

    public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Information;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
    {
        if (!IsEnabled(logLevel)) return;

        var message = formatter(state, exception);
        if (string.IsNullOrEmpty(message) && exception == null) return;

        var time = DateTime.Now.ToString("HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
        var level = GetShortLevel(logLevel);
        var shortCategory = _categoryName.Split('.').Last(); // Just get class name for brevity

        var logEntry = $"[{time}] [{level}] [{shortCategory}] {message}";
        if (exception != null)
        {
            logEntry += $" | Ex: {exception.Message}";
        }

        _provider.AddLog(logEntry);
    }

    private static string GetShortLevel(LogLevel level) => level switch
    {
        LogLevel.Trace => "TRC",
        LogLevel.Debug => "DBG",
        LogLevel.Information => "INF",
        LogLevel.Warning => "WRN",
        LogLevel.Error => "ERR",
        LogLevel.Critical => "CRT",
        _ => "UNK"
    };
}
