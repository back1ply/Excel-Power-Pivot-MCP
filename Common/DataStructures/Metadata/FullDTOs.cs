using System.Collections.Generic;

namespace ExcelPowerPivotMcp.Common.DataStructures.Metadata;

public class MeasureMetadata
{
    public required string Name { get; set; }
    public required string TableName { get; set; }
    public string? Caption { get; set; }
    public string? DataType { get; set; }
    public string? Expression { get; set; }
    public string? FormatString { get; set; }
    public string? Description { get; set; }
    public bool IsVisible { get; set; }
}

public class RelationshipMetadata
{
    public required string ForeignKeyTable { get; set; }
    public required string ForeignKeyColumn { get; set; }
    public required string PrimaryKeyTable { get; set; }
    public required string PrimaryKeyColumn { get; set; }
    public bool IsActive { get; set; }
}

public class ModelSummary
{
    public int TableCount { get; set; }
    public int MeasureCount { get; set; }
    public int RelationshipCount { get; set; }
    public long TotalRows { get; set; }
    public string? ConnectedWorkbook { get; set; }
    
    public List<TableMetadata> Tables { get; set; } = new();
    public List<MeasureMetadata> Measures { get; set; } = new();
    public List<Dictionary<string, string>> Relationships { get; set; } = new(); // Simplified view
    public List<Dictionary<string, object?>> TableRowCounts { get; set; } = new();
}

public class ColumnProfile
{
    public object? Min { get; set; }
    public object? Max { get; set; }
    public long DistinctCount { get; set; }
    public long Count { get; set; }
    public long NullCount { get; set; }
    public List<object?> DistinctValues { get; set; } = new();
    public string? Error { get; set; }
}

/// <summary>
/// Storage statistics for a table from DISCOVER_STORAGE_TABLES DMV.
/// Provides Vertipaq engine insights for performance optimization.
/// </summary>
public class StorageTableMetadata
{
    public required string TableName { get; set; }
    public string? TableId { get; set; }
    public long RowsCount { get; set; }
    public int PartitionCount { get; set; }
    public long? RiViolationCount { get; set; }
    public long? UsedSize { get; set; }
    public long? DataSize { get; set; }
}

/// <summary>
/// Storage statistics for a column from DISCOVER_STORAGE_TABLE_COLUMNS DMV.
/// Exposes Vertipaq compression, encoding, and memory usage details.
/// </summary>
public class StorageColumnMetadata
{
    public required string TableName { get; set; }
    public required string ColumnName { get; set; }
    public string? TableId { get; set; }
    public string? ColumnId { get; set; }
    public string? DataType { get; set; }
    public string? Encoding { get; set; }
    public long? DictionarySize { get; set; }
    public long? DataSize { get; set; }
    public long? Cardinality { get; set; }
    public bool? IsResidentInMemory { get; set; }
    public double? Temperature { get; set; }
    public DateTime? LastAccessed { get; set; }
}
