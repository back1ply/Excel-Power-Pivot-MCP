namespace ExcelPowerPivotMcp.Common.DataStructures;

public class RelationshipCreate
{
    public string ForeignTable { get; set; } = string.Empty;
    public string ForeignColumn { get; set; } = string.Empty;
    public string PrimaryTable { get; set; } = string.Empty;
    public string PrimaryColumn { get; set; } = string.Empty;
}
