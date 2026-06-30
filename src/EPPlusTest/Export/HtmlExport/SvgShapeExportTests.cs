using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class SvgShapeExportTests : TestBase
    {
        [TestMethod]
        public void ExportBasicShapeWorksheet()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("HtmlBasicSvgShape.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("ShapeWs");
                var rect = ws.Drawings.AddShape("SimpleRect", OfficeOpenXml.Drawing.eShapeStyle.Rect);
                //rect.Fill.Color = System.Drawing.Color.AliceBlue;
                //var rect2 = ws.Drawings.AddShape("SimpleRect2", OfficeOpenXml.Drawing.eShapeStyle.Rect);
                //rect2.Fill.Color = System.Drawing.Color.BlanchedAlmond;

                for (int i = 0; i< values.Count()*2; i++)
                {
                    if(i >= values.Count())
                    {
                        ws.Cells[i+1, 1].Value = -values[i - values.Count()];
                    }
                    else
                    {
                        ws.Cells[i+1, 1].Value = values[i];
                    }
                }

                var exporter = ws.Cells["A1:C20"].CreateHtmlExporter();

                exporter.Settings.Drawings.Include = ePictureInclude.IncludeInHtmlOnly;
                exporter.Settings.Drawings.DrawTypeInclude = eDrawingInclude.Shapes;

                var htmlPage = exporter.GetSinglePage();

                var file = GetOutputFile("html", "svgRect.html");

                SaveAndCleanup(package);

                File.WriteAllText(file.FullName, htmlPage);
            }
        }

        [TestMethod]
        public void ExportBarChart()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("HtmlSvgColChart.xlsx", true))
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

                var chart = ws.Drawings.AddBarChart("myColChart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
                chart.Series.Add(ws.Cells["A1:A8"]);

                //chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.BarChartStyle1);

                //chart.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style1;
                //var theme = ws.Workbook.ThemeManager.GetOrCreateTheme();
                //chart.StyleManager.SetChartStyle(0);
                //chart.StyleManager.SetChartStyle(0);
                //chart.StyleManager.ApplyStyles();
                //chart.Fill.Color = System.Drawing.Color.BlanchedAlmond;
                //chart.Series[0].Fill.Color = System.Drawing.Color.LightCoral;

                var exporter = ws.Cells["A1:C20"].CreateHtmlExporter();


                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Minify = false;
                exporter.Settings.Encoding = Encoding.UTF8;
                exporter.Settings.Drawings.Include = ePictureInclude.IncludeInHtmlOnly;
                exporter.Settings.Drawings.DrawTypeInclude = eDrawingInclude.Charts;

                var htmlPage = exporter.GetSinglePage();

                var file = GetOutputFile("html", "myColChart.html");
                var svgFile = GetOutputFile("html", "myColChartSvg.svg");

                File.WriteAllText(file.FullName, htmlPage);
                File.WriteAllText(svgFile.FullName, chart.ToSvg());

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ExportBarChartWithCategories()
        {
            int[] values = { 5, 10, 15, 20 };
            using (var package = OpenPackage("HtmlSvgColChartCategories.xlsx", true))
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

                var exporter = ws.Cells["A1:C20"].CreateHtmlExporter();

                exporter.Settings.Drawings.Include = ePictureInclude.IncludeInHtmlOnly;
                exporter.Settings.Drawings.DrawTypeInclude = eDrawingInclude.Charts;

                var htmlPage = exporter.GetSinglePage();

                var file = GetOutputFile("html", "colChartCats.html");

                SaveAndCleanup(package);

                File.WriteAllText(file.FullName, htmlPage);
            }
        }
    }
}
