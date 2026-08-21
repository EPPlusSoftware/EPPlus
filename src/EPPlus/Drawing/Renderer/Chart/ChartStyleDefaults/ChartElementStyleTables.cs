using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.Chart.ChartElementStyleTables
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

    internal static class ChartElementStyleTables
    {
        static Color GetLineColorForChartElement(ChartElement element, int ChartStyleId)
        {
            return Color.Empty;
            if(element.HasFlag(ChartElement.Axis | ChartElement.MajorGridLines))
            {
                if(ChartStyleId <= 32)
                {
                    //return Tx1
                }
                else
                {
                    //return dk1
                }
            }

        }
    }
}
