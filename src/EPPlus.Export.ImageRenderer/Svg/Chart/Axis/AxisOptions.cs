using OfficeOpenXml.Drawing.Chart;

internal class AxisOptions
{
    public double? LockedMin { get; set; }
    public double? LockedMax { get; set; }
    public double? LockedInterval { get; set; }
    public eTimeUnit? LockedIntervalUnit { get; set; }
    public bool AddPadding { get; set; } = false;
    public ExcelChartAxisStandard Axis { get; set; }
    public bool IsStacked100 { get; set; }
}