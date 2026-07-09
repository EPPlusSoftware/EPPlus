using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.Export.SvgExport
{
    [TestClass]
    public class EpplusGenerated : TestBase
    {
        [TestMethod]
        public void ColChartBlankCat()
        {
            int[] values = { 5, 10, 15, 20 };
            string name = "ColChartBlankCat";
            using (var package = OpenPackage($"{name}.xlsx", true))
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

                var svgFile = GetOutputFile("svg", $"{name}.svg");

                File.WriteAllText(svgFile.FullName, chart.ToSvg());

                var packageFile = GetOutputFile("svg", $"{name}.xlsx");
                package.SaveAs(packageFile);
            }
        }

        [TestMethod]
        public void ColChartCategories()
        {
            int[] values = { 5, 10, 15, 20 };
            string name = "ColChartCategories";
            using (var package = OpenPackage($"{name}.xlsx", true))
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

                var file = GetOutputFile("svg", $"{name}.svg");
                File.WriteAllText(file.FullName, chart.ToSvg());

                var packageFile = GetOutputFile("svg", $"{name}.xlsx");
                package.SaveAs(packageFile);
            }
        }


        [TestMethod]
        public void ColumnWithOneDtptPerSeries()
        {
            int[] values = { 5, 10, 15, 20 };
            string name = "ColumnOneDtPtPerSeries";
            using (var package = OpenPackage($"{name}.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");

                ws.Cells["A1"].Value = "Profit 2025";
                ws.Cells["B1"].Value = "Dates";

                for (int i = 0; i < values.Count() * 2; i++)
                {
                    if (i >= values.Count())
                    {
                        ws.Cells[i + 2, 1].Value = -values[i - values.Count()];
                        ws.Cells[i + 2, 2].Value = new DateTime(2026, 7, 1).AddDays(i);
                    }
                    else
                    {
                        ws.Cells[i + 2, 1].Value = values[i];
                        ws.Cells[i + 2, 2].Value = new DateTime(2026, 7, 1).AddDays(i);
                    }
                }

                ws.Cells["B2:B9"].Style.Numberformat.Format = "d-mmm-yy";

                var chart = ws.Drawings.AddBarChart("myColChart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
                chart.XAxis.AxisPosition = OfficeOpenXml.Drawing.Chart.eAxisPosition.Bottom;

                foreach (var cell in ws.Cells["A2:A9"])
                {
                    chart.Series.Add(cell);
                }

                chart.Fill.Color = System.Drawing.Color.BlanchedAlmond;
                chart.Series[0].Fill.Color = System.Drawing.Color.LightCoral;
                chart.SetPixelWidth(600);

                var file = GetOutputFile("svg", $"{name}.svg");

                File.WriteAllText(file.FullName, chart.ToSvg());

                var packageFile = GetOutputFile("svg", $"{name}.xlsx");

                package.SaveAs(packageFile);
            }
        }
    

    
        [TestMethod]
        public void ColumnWithSingleDtPt()
        {
            int[] values = { 5, 10, 15, 20 };
            string name = "singleDtPtColumn";

            using (var package = OpenPackage($"{name}.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");

                ws.Cells["A1"].Value = "Profit 2025";
                ws.Cells["B1"].Value = "Dates";

                for (int i = 0; i < values.Count() * 2; i++)
                {
                    if (i >= values.Count())
                    {
                        ws.Cells[i + 2, 1].Value = -values[i - values.Count()];
                        ws.Cells[i + 2, 2].Value = new DateTime(2026, 7, 1).AddDays(i);
                    }
                    else
                    {
                        ws.Cells[i + 2, 1].Value = values[i];
                        ws.Cells[i + 2, 2].Value = new DateTime(2026, 7, 1).AddDays(i);
                    }
                }

                ws.Cells["B2:B9"].Style.Numberformat.Format = "d-mmm-yy";

                var chart = ws.Drawings.AddBarChart("myColChart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
                chart.XAxis.AxisPosition = OfficeOpenXml.Drawing.Chart.eAxisPosition.Bottom;

                chart.Series.Add(ws.Cells["B2"]);


                chart.Fill.Color = System.Drawing.Color.BlanchedAlmond;
                chart.Series[0].Fill.Color = System.Drawing.Color.LightCoral;
                chart.SetPixelWidth(600);

                var file = GetOutputFile("svg", $"{name}.svg");

                File.WriteAllText(file.FullName, chart.ToSvg());

                var packageFile = GetOutputFile("svg", $"{name}.xlsx");

                package.SaveAs(packageFile);
            }
        }
    }
}