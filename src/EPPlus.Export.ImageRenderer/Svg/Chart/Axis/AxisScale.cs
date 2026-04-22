using EPPlus.Export.ImageRenderer;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;

internal class AxisScale
{
    public double Min { get; set; }
    public double Max { get; set; }
    public double MajorInterval { get; set; }
    public double MinorInterval { get; set; }
    public int TickCount { get; set; }
    public eTimeUnit? MajorDateUnit { get; set; }
    public eTimeUnit? MinorDateUnit { get; set; }
    public eTextOrientation TextOrientation { get; set; }
    public List<object> DisplayValues { get; set; }
}