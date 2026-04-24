namespace ExcelPowerPivotMcp.Common.DataStructures;

public class MeasureUpdate
{
    public string MeasureName { get; set; } = string.Empty;
    public string? NewName { get; set; }
    public string? Expression { get; set; }
    public string? Description { get; set; }
}
