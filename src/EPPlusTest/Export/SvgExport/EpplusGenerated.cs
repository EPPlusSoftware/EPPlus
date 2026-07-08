using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.Export.SvgExport
{
    [TestClass]
    public class EpplusGenerated : TestBase
    {
        [TestMethod]
        public void ColChartBlankCat()
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

        [TestMethod]
        public void ColChartCategories()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("Svg_GeneratedColChartCategories.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");

                ws.Cells["A1"].Value = "Value";
                ws.Cells["B1"].Value = "Title";

                for (int i = 0; i < values.Count() * 2; i++)
                {
                    if (i >= values.Count())
                    {
                        ws.Cells[i + 2, 1].Value = -values[i - values.Count()];
                    }
                    else
                    {
                        ws.Cells[i + 2, 1].Value = values[i];
                    }
                }

                var chart = ws.Drawings.AddBarChart("myColChart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
                chart.Series.Add(ws.Cells["A1:A10"]);

                chart.Fill.Color = System.Drawing.Color.BlanchedAlmond;
                chart.Series[0].Fill.Color = System.Drawing.Color.LightCoral;
                chart.SetPixelWidth(250);

                var file = GetOutputFile("svg", "colChartCategories.svg");
                File.WriteAllText(file.FullName, chart.ToSvg());


                SaveAndCleanup(package);
            }
        }
    }
}
