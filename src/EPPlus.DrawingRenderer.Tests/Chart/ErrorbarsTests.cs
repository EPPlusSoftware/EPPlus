using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class ErrorbarsTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForErrorbars_Sheet1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Errorbars.xlsx"))
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
                    SaveTextFileToWorkbook($"svg\\Errorbar_sheet1_{ix++}.svg", svg);
                }
            }
        }
    }
}
