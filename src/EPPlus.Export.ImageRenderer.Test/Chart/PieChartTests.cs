using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class PieChartTests : TestBase
    {

        [TestMethod]
        public void ReadAndGenerateExcelPieChartSvgs()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartSvgALL.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for(int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PieChartSvgALL\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }

        [TestMethod]
        public void ReadLegendIssue()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartSvgLegendIssue.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PieChartSvgLegendIssue\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }


        [TestMethod]
        public void ReadAndGenerateExcelPieChartPointExplosion()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PointExplosionStandAlone.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PointExplosionStandAlone\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }

        [TestMethod]
        public void BasicPieChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BasicPieChart.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var pChart = ws.Drawings[0].As.Chart.PieChart;

                var ser = pChart.Series[0];

                var legend = pChart.Legend;
                var entry = pChart.Legend.Entries;
                var pHeader = ser.Header;
                var pHeaderAddress = ser.HeaderAddress;
                var headerString = ser.GetHeaderString();


                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\BasicPieChart{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void Datalabels()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartDlblsOrig.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PieChartDlbls\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }

        [TestMethod]
        public void Datalabels2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartDlblsInsideEndOnly.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PieChartDlbls2\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }
    }
}
