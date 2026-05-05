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
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\BarChartForSvg_sheet1_{ix++}.svg", svg);
                }
            }
        }
    }
}
