using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace EPPlusTest.Export.SvgExport
{
    [TestClass]
    public class EpplusGenerated : TestBase
    {
        [TestMethod]
        public void GeneratedColChart()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("Svg_GeneratedColChart.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");

                for (int i = 0; i < values.Count() * 2; i++)
                {
                    if (i >= values.Count())
                    {
                        ws.Cells[i + 1, 1].Value = -values[i - values.Count()];
                    }
                    else
                    {
                        ws.Cells[i + 1, 1].Value = values[i];
                    }
                }

                var chart = ws.Drawings.AddBarChart("GeneratedColChart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
                chart.Series.Add(ws.Cells["A1:A8"]);

                var svgFile = GetOutputFile("svg", "GeneratedColChart.svg");

                File.WriteAllText(svgFile.FullName, chart.ToSvg());

                SaveAndCleanup(package);
            }
        }
    }
}
