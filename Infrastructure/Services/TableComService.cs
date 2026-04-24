using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.PowerPivot;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of Power Pivot table operations via COM.
/// All COM operations are:
/// - Dispatched to the STA thread via IExcelDispatcher
/// - Wrapped in ComScope for proper COM object cleanup
/// - Protected by ExecuteWithConnectionValidation for stale-connection detection
/// </summary>
public class TableComService : ITableComService
{
    private readonly IPowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;
    private readonly McpConfiguration _config;
    private readonly ILogger<TableComService> _logger;

    // Default timeout for COM operations (30 seconds)
    private static TimeSpan ComOperationTimeout => TimeSpan.FromSeconds(30);

    public TableComService(IPowerPivotConnection connection, IExcelDispatcher dispatcher, McpConfiguration config, ILogger<TableComService> logger)
    {
        _connection = connection;
        _dispatcher = dispatcher;
        _config = config;
        _logger = logger;
    }

    public async Task<string> AddTableToModelAsync(TableAddToModel def, CancellationToken ct = default)
    {
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);

                string actualTableName;
                if (def.UsePowerQuery)
                {
                    actualTableName = AddTableViaPowerQuery(workbook, model, def, scope);
                }
                else
                {
                    actualTableName = AddTableDirect(workbook, model, def, scope);
                }
                _connection.MarkDirty();
                return actualTableName;
            }, "adding table to model");
        }, ct);
    }

    private string AddTableViaPowerQuery(dynamic workbook, dynamic model, TableAddToModel def, ComScope scope)
    {
        // IMPORTANT: For Power Query tables, the Query name MUST match the desired ModelTable name
        // because the ModelTable name is automatically set to match the Query name and cannot be renamed.
        // See: https://techcommunity.microsoft.com/t5/excel/change-name-of-power-pivot-table-created-with-power-query/td-p/2721270

        // Validate QueryName if provided - must match TableName due to Excel COM API limitations
        if (!string.IsNullOrEmpty(def.QueryName) && !def.QueryName.Equals(def.TableName, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"QueryName '{def.QueryName}' must match TableName '{def.TableName}'. " +
                "Excel Power Pivot automatically names the ModelTable to match the Query name, " +
                "and the ModelTable.Name property is read-only (cannot be renamed). " +
                "Either omit the queryName parameter or set it equal to tableName.");
        }

        string queryName = def.TableName;

        dynamic queries = scope.TrackDynamic(workbook.Queries);

        if (ExcelComHelpers.ContainsName(queries, queryName))
            throw new ArgumentException($"A Power Query named '{queryName}' already exists.");

        dynamic modelTables = scope.TrackDynamic(model.ModelTables);
        if (ExcelComHelpers.ContainsName(modelTables, queryName))
            throw new ArgumentException($"Table '{queryName}' is already in the data model.");

        // Remember the count before adding
        int countBefore = modelTables.Count;

        string mFormula;
        if (!string.IsNullOrEmpty(def.MCode))
        {
            mFormula = def.MCode!;
        }
        else
        {
            // Verify table exists
            bool tableExists = false;
            dynamic worksheets = scope.TrackDynamic(workbook.Worksheets);
            ExcelComHelpers.EnumerateCollectionWithCleanup(worksheets, (Action<dynamic>)(sheet =>
            {
                if (tableExists) return;
                dynamic listObjects = sheet.ListObjects;
                if (ExcelComHelpers.ContainsName(listObjects, def.TableName))
                {
                    tableExists = true;
                }
                ComHelper.SafeRelease(listObjects);
            }));

            if (!tableExists) throw new ArgumentException($"Excel table '{def.TableName}' not found.");
            mFormula = $"let Source = Excel.CurrentWorkbook(){{[Name=\"{def.TableName}\"]}}[Content] in Source";
        }

        queries.Add(queryName, mFormula, $"Query to loading {def.TableName} to data model");

        dynamic connections = scope.TrackDynamic(workbook.Connections);
        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

        _logger.LogDebug("AddTableToModel Query: Name='Query - {QueryName}', ConnString='{ConnectionString}', Cmd='{Command}'",
            queryName, connectionString, queryName);

        connections.Add2(
            $"Query - {queryName}",
            $"Connection for Power Query {queryName}",
            connectionString,
            queryName,
            2,
            true,
            false
        );

        // Verify creation and get the actual table name created
        // Excel may have auto-renamed the table if there was a conflict
        string actualTableName = queryName; // Default fallback
        int countAfter = modelTables.Count;

        if (countAfter > countBefore)
        {
            // Find the newly added table (should be the last one, but let's verify)
            try
            {
                dynamic newTable = modelTables[countAfter];
                actualTableName = (string)newTable.Name;
                ComHelper.SafeRelease(newTable);
            }
            catch
            {
                // Fallback: enumerate to find new table
                var foundNames = new List<string>();
                ExcelComHelpers.EnumerateCollectionWithCleanup(modelTables, (Action<dynamic>)(table =>
                {
                    string tableName = (string)table.Name;
                    // Look for a table that starts with our requested name
                    if (tableName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        tableName.StartsWith(queryName, StringComparison.OrdinalIgnoreCase))
                    {
                        foundNames.Add(tableName);
                    }
                }));

                if (foundNames.Count > 0)
                {
                    actualTableName = foundNames[0];
                }
            }
        }

        if (!actualTableName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("Power Query table created with name '{ActualName}' instead of requested '{RequestedName}'",
                actualTableName, queryName);
        }

        _logger.LogInformation("Created Power Query table '{TableName}' in data model", actualTableName);
        return actualTableName;
    }

    private string AddTableDirect(dynamic workbook, dynamic model, TableAddToModel def, ComScope scope)
    {
        dynamic modelTables = scope.TrackDynamic(model.ModelTables);
        if (ExcelComHelpers.ContainsName(modelTables, def.TableName))
            throw new ArgumentException($"Table '{def.TableName}' is already in the data model.");

        bool tableExists = false;
        dynamic worksheets = scope.TrackDynamic(workbook.Worksheets);
        ExcelComHelpers.EnumerateCollectionWithCleanup(worksheets, (Action<dynamic>)(sheet =>
        {
            if (tableExists) return;
            dynamic listObjects = sheet.ListObjects;
            if (ExcelComHelpers.ContainsName(listObjects, def.TableName))
            {
                tableExists = true;
            }
            ComHelper.SafeRelease(listObjects);
        }));

        if (!tableExists) throw new ArgumentException($"Excel table '{def.TableName}' not found.");

        dynamic connections = scope.TrackDynamic(workbook.Connections);
        string connectionString = "WORKSHEET;";
        string commandText = def.TableName;

        _logger.LogDebug("AddTableToModel Direct: Name='WorksheetConnection_{TableName}', ConnString='{ConnectionString}', Cmd='{CommandText}'",
            def.TableName, connectionString, commandText);

        connections.Add2(
            $"WorksheetConnection_{def.TableName}",
            $"Connection to Excel table {def.TableName}",
            connectionString,
            commandText,
            7,
            true,
            false
        );

        // NOTE: The ModelTable.Name property is READ-ONLY for all connection types per Excel COM API.
        // For linked tables, the ModelTable inherits the name from the Excel table.
        // If the Excel table is renamed later, the Power Pivot table name must be updated manually.
        // See: https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-the-powerpivot-model-object-in-excel
        _logger.LogInformation("Created linked table '{TableName}' in data model", def.TableName);
        return def.TableName;
    }

    public async Task DeleteTableAsync(string tableName, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelTables = scope.TrackDynamic(model.ModelTables);

                dynamic table = ExcelComHelpers.FindByNameOrThrow(modelTables, tableName, "Table");
                scope.TrackDynamic(table);

                // CRITICAL: To delete a table from the Data Model, we must delete its Source Connection.
                // Just calling table.Delete() often fails or leaves the connection orphaned.
                bool connectionDeleted = false;
                try
                {
                    dynamic conn = table.SourceWorkbookConnection;
                    if (conn != null)
                    {
                        scope.TrackDynamic(conn);
                        string connName = conn.Name; // Convert dynamic to string first
                        _logger.LogInformation("Deleting source connection '{ConnectionName}' to remove table '{TableName}'", connName, tableName);
                        conn.Delete();
                        connectionDeleted = true;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not delete source connection for table '{TableName}'", tableName);
                }

                // Fallback: If no connection was found/deleted, try deleting the table directly.
                if (!connectionDeleted)
                {
                    _logger.LogInformation("Attempting direct delete of ModelTable '{TableName}'", tableName);
                    table.Delete();
                }

                _connection.MarkDirty();
            }, "deleting table");
        }, ct);
    }

    public async Task<List<Dictionary<string, object?>>> GetPowerQueriesAsync(CancellationToken ct = default)
    {
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic queries = scope.TrackDynamic(workbook.Queries);

                return ExcelComHelpers.EnumerateCollectionWithCleanup(queries, (Func<dynamic, Dictionary<string, object?>>)(q =>
                {
                    return new Dictionary<string, object?>
                    {
                        ["name"] = (string)q.Name,
                        ["formula"] = (string)q.Formula,
                        ["description"] = q.Description != null ? (string)q.Description : null
                    };
                }));
            }, "retrieving Power Queries");
        }, ct);
    }

    public async Task<List<ColumnMetadata>> GetCalculatedColumnsAsync(string? tableName = null, CancellationToken ct = default)
    {
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();
                var columns = new List<ColumnMetadata>();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelTables = scope.TrackDynamic(model.ModelTables);

                ExcelComHelpers.EnumerateCollectionWithCleanup(modelTables, (Action<dynamic>)(table =>
                {
                    string tblName = table.Name;
                    if (!string.IsNullOrEmpty(tableName) && !tblName.Equals(tableName, StringComparison.OrdinalIgnoreCase)) return;

                    dynamic tableColumns = table.ModelTableColumns;
                    ExcelComHelpers.EnumerateCollectionWithCleanup(tableColumns, (Action<dynamic>)(col =>
                    {
                        try
                        {
                            string? formula = null;
                            try { formula = col.Formula; }
                            catch { return; } // Column doesn't have Formula property (not calculated)

                            if (!string.IsNullOrEmpty(formula))
                            {
                                columns.Add(new ColumnMetadata
                                {
                                    Name = (string)col.Name,
                                    TableName = tblName,
                                    IsCalculated = true,
                                    Expression = formula,
                                    DataType = col.DataType.ToString()
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Skipped column due to error: {ex.Message}");
                        }
                    }));
                    ComHelper.SafeRelease(tableColumns);
                }));

                return columns;
            }, "retrieving calculated columns");
        }, ct);
    }

    public async Task RefreshTableAsync(string tableName, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelTables = scope.TrackDynamic(model.ModelTables);

                dynamic table = ExcelComHelpers.FindByNameOrThrow(modelTables, tableName, "Table");
                scope.TrackDynamic(table);
                table.Refresh();
            }, "refreshing table");
        }, ct);
    }
}
