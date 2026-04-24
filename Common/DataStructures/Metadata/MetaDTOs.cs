using System.Collections.Generic;

namespace ExcelPowerPivotMcp.Common.DataStructures.Metadata;

public class TableMetadata
{
    public required string Name { get; set; }
    public string? Caption { get; set; }
    public bool IsVisible { get; set; }
    public string? Description { get; set; }
    public long? RowCount { get; set; }
}

public class ColumnMetadata
{
    public required string Name { get; set; }
    public required string TableName { get; set; }
    public string? Caption { get; set; }
    public string? DataType { get; set; }
    public string? Description { get; set; }
    public bool IsHidden { get; set; }
    
    // For Calculated Columns
    public bool IsCalculated { get; set; }
    public string? Expression { get; set; }
}
