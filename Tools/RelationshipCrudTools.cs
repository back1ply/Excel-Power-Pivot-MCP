using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;
using ExcelPowerPivotMcp.Common;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for relationship management.
/// </summary>
[McpServerToolType]
public class RelationshipCrudTools
{
    private readonly Core.Services.IRelationshipService _relationshipService;

    public RelationshipCrudTools(Core.Services.IRelationshipService relationshipService)
    {
        _relationshipService = relationshipService;
    }

    [McpServerTool(Name = "create_relationship")]
    [Description("Create a one-to-many relationship between tables. Requires confirm=true.")]
    public async Task<string> CreateRelationship(
        [Description("Table containing the foreign key (many side)")] string foreignTable,
        [Description("Column in the foreign key table")] string foreignColumn,
        [Description("Table containing the primary key (one side)")] string primaryTable,
        [Description("Column in the primary key table")] string primaryColumn,
        [Description("Set to true to confirm creation. Required for safety.")] bool confirm = false)
    {
        if (string.IsNullOrWhiteSpace(foreignTable))
            throw new ArgumentException("foreign_table is required");
        if (string.IsNullOrWhiteSpace(foreignColumn))
            throw new ArgumentException("foreign_column is required");
        if (string.IsNullOrWhiteSpace(primaryTable))
            throw new ArgumentException("primary_table is required");
        if (string.IsNullOrWhiteSpace(primaryColumn))
            throw new ArgumentException("primary_column is required");
            
        if (!confirm)
        {
            return JsonSerializer.Serialize(new
            {
                requiresConfirmation = true,
                message = $"This will create a relationship: {foreignTable}.{foreignColumn} → {primaryTable}.{primaryColumn}. Call again with confirm=true to proceed.",
                warning = "Creating relationships alters the data model structure. Verify table/column names with list_tables and list_columns first."
            });
        }
        
        var request = new Common.DataStructures.RelationshipCreate
        {
            ForeignTable = foreignTable,
            ForeignColumn = foreignColumn,
            PrimaryTable = primaryTable,
            PrimaryColumn = primaryColumn
        };

        await _relationshipService.CreateRelationshipAsync(request);

        var result = new
        {
            success = true,
            message = $"Successfully created relationship: {foreignTable}.{foreignColumn} → {primaryTable}.{primaryColumn}. Remember to call save_workbook to persist changes.",
            data = request
        };
        return JsonSerializer.Serialize(result);
    }

    [McpServerTool(Name = "delete_relationship")]
    [Description("Delete a relationship between tables. Requires confirm=true.")]
    public async Task<string> DeleteRelationship(
        [Description("Table containing the foreign key")] string foreignTable,
        [Description("Column in the foreign key table")] string foreignColumn,
        [Description("Table containing the primary key")] string primaryTable,
        [Description("Column in the primary key table")] string primaryColumn,
        [Description("Set to true to confirm deletion. Required for safety.")] bool confirm = false)
    {
        if (string.IsNullOrWhiteSpace(foreignTable))
            throw new ArgumentException("foreign_table is required");
        if (string.IsNullOrWhiteSpace(foreignColumn))
            throw new ArgumentException("foreign_column is required");
        if (string.IsNullOrWhiteSpace(primaryTable))
            throw new ArgumentException("primary_table is required");
        if (string.IsNullOrWhiteSpace(primaryColumn))
            throw new ArgumentException("primary_column is required");
            
        if (!confirm)
        {
            return JsonSerializer.Serialize(new
            {
                requiresConfirmation = true,
                message = $"This will permanently delete relationship: {foreignTable}.{foreignColumn} → {primaryTable}.{primaryColumn}. Call again with confirm=true to proceed.",
                warning = "Deleting relationships may break DAX measures that rely on table traversal!"
            });
        }
        
        await _relationshipService.DeleteRelationshipAsync(
            foreignTable, foreignColumn, primaryTable, primaryColumn);

        var result = new
        {
            success = true,
            message = $"Successfully deleted relationship: {foreignTable}.{foreignColumn} → {primaryTable}.{primaryColumn}. Remember to call save_workbook to persist changes."
        };
        return JsonSerializer.Serialize(result);
    }

    [McpServerTool(Name = "set_relationship_active")]
    [Description("Activate or deactivate a relationship.")]
    public async Task<string> SetRelationshipActive(
        [Description("Table containing the foreign key")] string foreignTable,
        [Description("Column in the foreign key table")] string foreignColumn,
        [Description("Table containing the primary key")] string primaryTable,
        [Description("Column in the primary key table")] string primaryColumn,
        [Description("True to activate, false to deactivate")] bool active)
    {
        await _relationshipService.SetRelationshipActiveAsync(
            foreignTable, foreignColumn, primaryTable, primaryColumn, active);

        var result = new
        {
            success = true,
            message = $"Successfully {(active ? "activated" : "deactivated")} relationship: {foreignTable}.{foreignColumn} → {primaryTable}.{primaryColumn}. Remember to call save_workbook to persist changes."
        };
        return JsonSerializer.Serialize(result);
    }
}
