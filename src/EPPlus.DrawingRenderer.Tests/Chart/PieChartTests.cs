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
        public void ReadAndCreateSvgsAll()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartSvgALL.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
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
        public void BestFitPie()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            //Same as PieChartSvgALL but seperatated out explosionBestFit chart
            using (var p = OpenTemplatePackage("BestFitPie.xlsx"))
            {
                var ws = p.Workbook.Worksheets["bestFit"];

                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\PieChartSvgALL\\ExplosionOwnSheet_{ws.Name}_{c.Name}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void SimpleBestFit()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SimpleBestFit.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }
        [TestMethod]
        public void GenerateEPPlusPieCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("2.4-CreateAFileSystemReport.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Pie{i}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateEPPlusPieChartsErrorCase()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("2.4-CreateAFileSystemReport_errorpie.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\PieError{i}.svg", svg);
                }
            }
        }
    }
}
