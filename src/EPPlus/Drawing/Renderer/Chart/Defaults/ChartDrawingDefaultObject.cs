using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Encryption;
using System;
using System.Drawing;
using tc = OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing.Renderer.Chart.Defaults
{
    [Flags]
    enum ChartElement
    {
        None = 0,
        ChartArea = 1,
        PlotArea2d = 2,
        PloatArea3d = 4,
        Axis = 8,
        MinorGridLines = 16,
        MajorGridLines = 32,
        DataTable = 64,
        Floor = 128,
        Walls = 256,
        OtherLines = 512,
    }
}
