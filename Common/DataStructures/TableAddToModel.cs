namespace ExcelPowerPivotMcp.Common.DataStructures;

public class TableAddToModel
{
    public string TableName { get; set; } = string.Empty;
    
    // For Power Query
    public bool UsePowerQuery { get; set; }
    public string? QueryName { get; set; }
    public string? MCode { get; set; }
}
