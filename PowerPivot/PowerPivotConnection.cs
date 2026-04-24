using System;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelPowerPivotMcp.Common.Utils;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Represents an Excel workbook with a Power Pivot model
/// </summary>
public class PowerPivotWorkbook
{
    public string Name { get; }
    public string FullPath { get; }
    public bool HasDataModel { get; }
    public int ProcessId { get; }
    
    public PowerPivotWorkbook(string name, string fullPath, bool hasDataModel, int processId = 0)
    {
        Name = name;
        FullPath = fullPath;
        HasDataModel = hasDataModel;
        ProcessId = processId;
    }
}

/// <summary>
/// Manages connection to Power Pivot data models via Excel's ADO connection.
///
/// This is a partial class split across multiple files for maintainability:
/// - PowerPivotConnection.cs (this file): Core connection management
/// - PowerPivotConnection.Discovery.cs: Workbook discovery
/// - PowerPivotConnection.Measures.cs: Measure CRUD operations
/// - PowerPivotConnection.Relationships.cs: Relationship operations
/// - PowerPivotConnection.Tables.cs: Table and Power Query operations
/// - PowerPivotConnection.Query.cs: DAX execution and metadata queries
/// </summary>
public partial class PowerPivotConnection : IPowerPivotConnection
{
    #region Singleton & Fields

    private static PowerPivotConnection? _instance;
    private static readonly object _instanceLock = new();
    private static ILogger? _logger;

    /// <summary>
    /// Set the logger for PowerPivotConnection (should be called during application startup).
    /// </summary>
    public static void SetLogger(ILogger<PowerPivotConnection> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Thread-safe resettable singleton instance.
    /// Use Reset() after detecting fatal COM errors to allow reconnection.
    /// </summary>
    public static PowerPivotConnection Instance
    {
        get
        {
            lock (_instanceLock)
            {
                _instance ??= new PowerPivotConnection();
                return _instance;
            }
        }
    }

    /// <summary>
    /// Reset the singleton instance after a fatal COM error.
    /// This allows the server to recover from Excel crashes without restarting.
    /// </summary>
    public static void Reset()
    {
        lock (_instanceLock)
        {
            if (_instance != null)
            {
                _logger?.LogWarning("Resetting singleton instance due to fatal COM error");
                _instance.Disconnect();
                _instance = null;
            }
        }
    }

    /// <summary>
    /// Check if a COMException indicates a fatal/unrecoverable error
    /// that requires resetting the connection.
    /// </summary>
    public static bool IsFatalComError(System.Runtime.InteropServices.COMException ex)
    {
        // Fatal HRESULTs that require full reset:
        // RPC_E_DISCONNECTED (0x80010108) - The object invoked has disconnected from its clients
        // CO_E_OBJNOTCONNECTED (0x800401FD) - Object is not connected to server
        // RPC_E_SERVER_DIED (0x80010007) - The callee (server) process exited unexpectedly
        // RPC_E_SERVER_DIED_DNE (0x8001010D) - The callee died; could not dispatch
        // RPC_S_SERVER_UNAVAILABLE (0x800706BA) - RPC server unavailable
        // RPC_S_CALL_FAILED (0x800706BE) - RPC call failed
        return ex.HResult == unchecked((int)0x80010108) || // RPC_E_DISCONNECTED
               ex.HResult == unchecked((int)0x800401FD) || // CO_E_OBJNOTCONNECTED
               ex.HResult == unchecked((int)0x80010007) || // RPC_E_SERVER_DIED
               ex.HResult == unchecked((int)0x8001010D) || // RPC_E_SERVER_DIED_DNE
               ex.HResult == unchecked((int)0x800706BA) || // RPC_S_SERVER_UNAVAILABLE
               ex.HResult == unchecked((int)0x800706BE);   // RPC_S_CALL_FAILED
    }

    private string? _connectedWorkbookPath;
    private object? _adoConnection; // Excel's ADO connection object
    private object? _workbook; // Reference to the connected workbook
    private readonly object _lock = new();

    // P/Invoke for GetActiveObject (not available in .NET Core)
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);


    public bool IsConnected => _adoConnection != null && _connectedWorkbookPath != null;
    public string? ConnectedWorkbook => _connectedWorkbookPath;
    
    /// <summary>
    /// Expose the underlying COM Workbook object for Core operations.
    /// </summary>
    public dynamic? Workbook => _workbook;

    /// <summary>
    /// Expose the underlying ADO Connection object for Core operations.
    /// </summary>
    public dynamic? AdoConnection => _adoConnection;

    /// <summary>
    /// Track whether the model has unsaved changes.
    /// Set to true by CRUD operations, cleared by SaveWorkbook().
    /// </summary>
    public bool IsDirty { get; private set; }

    /// <summary>
    /// Mark the model as having unsaved changes.
    /// Call this after any mutation operation (create/update/delete measure, relationship, etc.)
    /// </summary>
    public void MarkDirty() => IsDirty = true;

    /// <summary>
    /// Clear the dirty flag after saving.
    /// </summary>
    public void ClearDirty() => IsDirty = false;
    
    #endregion

    #region Connection Management

    /// <summary>
    /// Connect to a Power Pivot model in a workbook using Excel's ADO connection.
    /// This method performs a single connection attempt. 
    /// Retry logic is handled by the caller (ExcelInteropService) to ensure thread safety.
    /// </summary>
    public void Connect(string workbookPath)
    {
        lock (_lock)
        {
            Disconnect();

            var excelApp = GetExcelApplication();
            if (excelApp == null)
            {
                throw new InvalidOperationException("Excel is not running. Please open the workbook in Excel first.");
            }

            try
            {
                using var scope = new ComScope();
                scope.TrackDynamic(excelApp);
                
                dynamic dynExcel = excelApp;
                var workbooksCollection = scope.TrackDynamic(dynExcel.Workbooks);

                if (workbooksCollection == null)
                {
                    throw new InvalidOperationException("Cannot access Excel workbooks.");
                }

                string normalizedPath = workbookPath.ToLowerInvariant();

                // Find the workbook by path
                object? targetWorkbook = null;
                int count = workbooksCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = workbooksCollection.Item[i];
                    bool isMatch = false;
                    try 
                    {
                        string fullPath = wb.FullName;
                        if (fullPath.Equals(normalizedPath, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWorkbook = wb;
                            isMatch = true;
                            // Do NOT release 'wb' if it matches, we keep it as targetWorkbook
                            break; 
                        }
                    }
                    finally
                    {
                        if (!isMatch)
                        {
                            ComHelper.SafeRelease(wb);
                        }
                    }
                }

                if (targetWorkbook == null)
                {
                    throw new InvalidOperationException($"Workbook not found in Excel: {workbookPath}. Make sure it's open.");
                }

                dynamic dynWorkbook = targetWorkbook;
                
                // Get the Model object
                var model = dynWorkbook.Model;
                if (model == null)
                {
                    // If model access fails (e.g. not initialized), release workbook
                    ComHelper.SafeRelease(targetWorkbook);
                    throw new InvalidOperationException("Workbook does not have a Power Pivot data model.");
                }

                // Get DataModelConnection from the Model
                var dataModelConnection = model.DataModelConnection;
                if (dataModelConnection == null)
                {
                    ComHelper.SafeRelease(targetWorkbook);
                    throw new InvalidOperationException("Cannot access DataModelConnection.");
                }

                // Get ModelConnection
                var modelConnection = dataModelConnection.ModelConnection;
                if (modelConnection == null)
                {
                    ComHelper.SafeRelease(targetWorkbook);
                    throw new InvalidOperationException("Cannot access ModelConnection.");
                }

                // Get ADOConnection - this gives us a live connection to the data model
                _adoConnection = modelConnection.ADOConnection;
                if (_adoConnection == null)
                {
                    ComHelper.SafeRelease(targetWorkbook);
                    throw new InvalidOperationException("Cannot get ADO connection from the data model.");
                }

                // IMPORTANT: Force model load by executing a dummy command.
                // This warmup operation adds ~2 seconds to initial connection time but is essential
                // for reliable operation. Without it, the first DAX query or metadata operation may
                // fail or timeout. The warmup ensures Excel has fully initialized the data model
                // and the ADO connection is ready for use.
                //
                // Connection timeline (observed in production):
                // - Retry mechanism (if first attempt fails): +1-2 seconds
                // - Data model warmup (this operation): ~2 seconds per warmup
                // - Total initial connection time: typically 4-6 seconds
                try
                {
                    _logger?.LogInformation("Warming up data model...");
                    dynamic dynAdoConn = _adoConnection;
                    // Open connection if closed
                    if (dynAdoConn.State == 0) dynAdoConn.Open();
                    // Execute a simple DAX query to ensure the data model is loaded
                    dynAdoConn.Execute("EVALUATE ROW(\"Warmup\", 1)");
                    _logger?.LogInformation("Data model warmed up successfully");
                }
                catch (Exception wEx)
                {
                    _logger?.LogWarning(wEx, "Warm-up warning (non-fatal, connection may still work)");
                }

                _workbook = targetWorkbook;
                _connectedWorkbookPath = workbookPath;
            }
            catch (Exception ex) when (ex is not InvalidOperationException)
            {
                throw new InvalidOperationException($"Error connecting to data model: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// Disconnect from current Power Pivot model
    /// </summary>
    public void Disconnect()
    {
        lock (_lock)
        {
            ComHelper.SafeRelease(_adoConnection);
            _adoConnection = null;
            
            ComHelper.SafeRelease(_workbook);
            _workbook = null;
            
            _connectedWorkbookPath = null;
        }
    }

    /// <summary>
    /// Check if the current connection is still valid (Excel and workbook still open).
    /// </summary>
    /// <returns>True if connected and the workbook is still accessible.</returns>
    public bool IsConnectionValid()
    {
        if (!IsConnected || _workbook == null) return false;
        
        try
        {
            // Quick check: try to read the workbook name
            dynamic wb = _workbook;
            _ = wb.Name;
            return true;
        }
        catch (System.Runtime.InteropServices.COMException ex) when (IsFatalComError(ex))
        {
            // COM object is stale - Excel or workbook closed
            _logger?.LogWarning(ex, "Connection lost (HRESULT: 0x{HResult:X8}) - Excel or workbook was closed", ex.HResult);
            Reset(); // Full reset to allow reconnection
            return false;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            // Other COM error - log but don't disconnect
            _logger?.LogWarning(ex, "COM error during validation");
            return false;
        }
        catch
        {
            return false;
        }
    }

    // Note: IsStaleConnectionError was consolidated into the public IsFatalComError method above

    /// <summary>
    /// Ensure we have an active connection to a workbook
    /// </summary>
    public void EnsureConnected()
    {
        if (!IsConnected || _workbook == null)
        {
            throw new InvalidOperationException("Not connected to a workbook. Call connect_workbook first.");
        }
        
        if (!IsConnectionValid())
        {
            throw new InvalidOperationException(
                "Connection to workbook lost. Excel may have been closed or the workbook was closed. " +
                "Use discover_workbooks to find available workbooks and connect_workbook to reconnect.");
        }
    }

    /// <summary>
    /// Execute an operation with proactive stale-connection detection.
    /// Catches COM exceptions during execution and provides actionable error messages.
    /// </summary>
    /// <typeparam name="T">Return type</typeparam>
    /// <param name="operation">The operation to execute</param>
    /// <param name="operationName">Description of the operation for error messages</param>
    /// <returns>The operation result</returns>
    public T ExecuteWithConnectionValidation<T>(Func<T> operation, string operationName)
    {
        EnsureConnected();
        
        try
        {
            return operation();
        }
        catch (System.Runtime.InteropServices.COMException ex) when (IsFatalComError(ex))
        {
            _logger?.LogError(ex, "Connection lost during {OperationName} (HRESULT: 0x{HResult:X8})", operationName, ex.HResult);
            Reset(); // Full reset to allow reconnection
            throw new InvalidOperationException(
                $"Connection lost during {operationName}. Excel may have crashed or the workbook was closed. " +
                "Use discover_workbooks and connect_workbook to reconnect.", ex);
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            // Detect duplicate measure/relationship errors and provide helpful guidance
            if (ex.Message.Contains("CalculationProperty") && ex.Message.Contains("key"))
            {
                // Extract the object name from the error message (between single quotes)
                string objectName = "the object";
                var match = System.Text.RegularExpressions.Regex.Match(ex.Message, @"\[([^\]]+)\]");
                if (match.Success)
                {
                    objectName = match.Groups[1].Value;
                }

                string suggestion = operationName switch
                {
                    "creating measure" => $"A measure named '{objectName}' already exists. Use update_measure to modify it instead of create_measure.",
                    _ => $"An object named '{objectName}' already exists in the model."
                };

                throw new InvalidOperationException(
                    $"Duplicate object error during {operationName}. {suggestion} Details: {ex.Message}", ex);
            }

            throw new InvalidOperationException(
                $"COM error during {operationName}. Excel may be busy or showing a dialog. Details: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Execute an operation with proactive stale-connection detection (void version).
    /// </summary>
    /// <param name="action">The action to execute</param>
    /// <param name="operationName">Description of the operation for error messages</param>
    public void ExecuteWithConnectionValidation(Action action, string operationName)
    {
        ExecuteWithConnectionValidation(() => { action(); return true; }, operationName);
    }

    /// <summary>
    /// Save the connected workbook to persist any changes made to the data model.
    /// </summary>
    public void SaveWorkbook()
    {
        ExecuteWithConnectionValidation(() =>
        {
            dynamic dynWorkbook = _workbook!;
            dynWorkbook.Save();
            ClearDirty(); // Clear dirty flag after successful save
        }, "saving workbook");
    }
    
    #endregion
}
