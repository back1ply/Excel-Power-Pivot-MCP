using System;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Interface for Power Pivot connection management.
/// Provides core functionality for infrastructure services to interact with Excel Power Pivot models.
/// </summary>
public interface IPowerPivotConnection
{
    /// <summary>
    /// True if connected to a workbook with a data model.
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// Path to the currently connected workbook, or null if not connected.
    /// </summary>
    string? ConnectedWorkbook { get; }

    /// <summary>
    /// True if the model has unsaved changes.
    /// </summary>
    bool IsDirty { get; }

    /// <summary>
    /// Expose the underlying COM Workbook object for service operations.
    /// </summary>
    dynamic? Workbook { get; }

    /// <summary>
    /// Expose the underlying ADO Connection object for DMV and DAX queries.
    /// </summary>
    dynamic? AdoConnection { get; }

    /// <summary>
    /// Mark the model as having unsaved changes.
    /// Call this after any mutation operation (create/update/delete measure, relationship, etc.)
    /// </summary>
    void MarkDirty();

    /// <summary>
    /// Clear the dirty flag after saving.
    /// </summary>
    void ClearDirty();

    /// <summary>
    /// Ensure we have an active connection to a workbook.
    /// Throws InvalidOperationException if not connected or connection is stale.
    /// </summary>
    void EnsureConnected();

    /// <summary>
    /// Check if the current connection is still valid (Excel and workbook still open).
    /// </summary>
    /// <returns>True if connected and the workbook is still accessible.</returns>
    bool IsConnectionValid();

    /// <summary>
    /// Execute an operation with proactive stale-connection detection.
    /// Catches COM exceptions during execution and provides actionable error messages.
    /// </summary>
    /// <typeparam name="T">Return type</typeparam>
    /// <param name="operation">The operation to execute</param>
    /// <param name="operationName">Description of the operation for error messages</param>
    /// <returns>The operation result</returns>
    T ExecuteWithConnectionValidation<T>(Func<T> operation, string operationName);

    /// <summary>
    /// Execute an operation with proactive stale-connection detection (void version).
    /// </summary>
    /// <param name="action">The action to execute</param>
    /// <param name="operationName">Description of the operation for error messages</param>
    void ExecuteWithConnectionValidation(Action action, string operationName);

    /// <summary>
    /// Save the connected workbook to persist any changes made to the data model.
    /// </summary>
    void SaveWorkbook();
}
