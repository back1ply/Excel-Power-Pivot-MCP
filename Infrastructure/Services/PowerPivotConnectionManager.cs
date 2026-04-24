using ExcelPowerPivotMcp.PowerPivot;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// DI-friendly wrapper around the PowerPivotConnection singleton.
///
/// This solves the singleton/DI mismatch issue: when PowerPivotConnection.Reset()
/// is called after a fatal COM error, the old singleton instance is nulled.
/// Services that captured the old instance at DI resolution time would hold stale references.
///
/// By using this manager (which always delegates to Instance), all services automatically
/// get the current valid connection after a reset.
/// </summary>
public class PowerPivotConnectionManager : IPowerPivotConnectionManager
{
    /// <summary>
    /// Constructor that initializes the static logger in PowerPivotConnection.
    /// This is required because PowerPivotConnection is a singleton accessed before full DI resolution.
    /// </summary>
    public PowerPivotConnectionManager(ILogger<PowerPivotConnection> logger)
    {
        PowerPivotConnection.SetLogger(logger);
    }

    /// <inheritdoc/>
    public PowerPivotConnection Current => PowerPivotConnection.Instance;

    /// <inheritdoc/>
    public bool IsConnected => Current.IsConnected;

    /// <inheritdoc/>
    public string? ConnectedWorkbook => Current.ConnectedWorkbook;

    /// <inheritdoc/>
    public void Reset() => PowerPivotConnection.Reset();

    /// <summary>
    /// Check if a COMException indicates a fatal/unrecoverable error.
    /// Static helper, does not belong to the interface.
    /// </summary>
    public static bool IsFatalComError(System.Runtime.InteropServices.COMException ex)
        => PowerPivotConnection.IsFatalComError(ex);
}
