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
            using (var p = OpenTemplatePackage("PieChartSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                for(int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = renderer.RenderDrawingToSvg(c);
                        SaveTextFileToWorkbook($"svg\\PieChartSvg\\s{i}_{ws.Name}_{c.Name}.svg", svg);
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
                        SaveTextFileToWorkbook($"svg\\PieChartDlbls\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }
    }
}
