using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class TrendlineTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForTrendlines_Sheet1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Trendlines.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Trendline_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Trendline_sheet1_{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void TrendlineAlt()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("TrendlineAlt.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Trendline_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\TrendlineAlt_sheet1_{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForTrendlines_Sheet2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Trendlines.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                //var ix = 3;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Trendline_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Trendline_sheet2_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForTrendlineCharts_sheet5()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[4];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\Trendline_sheet5_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Trendline_Sheet5{ix++}.svg", svg);
                }
            }
        }

    }
}
