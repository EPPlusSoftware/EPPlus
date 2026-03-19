using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;

internal class AxisOptions
{
    public double? LockedMin { get; set; }
    public double? LockedMax { get; set; }
    public double? LockedInterval { get; set; }
    public eTimeUnit? LockedIntervalUnit { get; set; }
    public bool AddPadding { get; set; } = false;
    public ExcelChartAxisStandard Axis { get; set; }
    public bool IsStacked100 { get; set; }
    public RenderItem ChartSize { get; set; }
    public string NumberFormat { get; set; }
}