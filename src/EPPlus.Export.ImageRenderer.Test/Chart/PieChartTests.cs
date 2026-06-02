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

                for(int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\PieChartSvgALL\\s{i}_{ws.Name}_{c.Name}.svg", svg);
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

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
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

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
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

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
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

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\PieChartDlbls\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }
    }
}
