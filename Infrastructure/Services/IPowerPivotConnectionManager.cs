using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Interface for managing the PowerPivot connection.
/// Abstraction allows for mocking connection state in unit tests.
/// </summary>
public interface IPowerPivotConnectionManager
{
    /// <summary>
    /// Get the current PowerPivotConnection instance.
    /// </summary>
    PowerPivotConnection Current { get; }

    /// <summary>
    /// Check if the current connection is valid and connected.
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Get the path of the currently connected workbook.
    /// </summary>
    string? ConnectedWorkbook { get; }

    /// <summary>
    /// Reset the connection after a fatal COM error.
    /// </summary>
    void Reset();
}
