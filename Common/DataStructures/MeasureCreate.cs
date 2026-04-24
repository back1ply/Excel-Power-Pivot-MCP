namespace ExcelPowerPivotMcp.Common.DataStructures;

public class MeasureCreate
{
    public string TableName { get; set; } = string.Empty;
    public string MeasureName { get; set; } = string.Empty;
    public string Expression { get; set; } = string.Empty;
    public string? Description { get; set; }
    public string? FormatString { get; set; }
}
