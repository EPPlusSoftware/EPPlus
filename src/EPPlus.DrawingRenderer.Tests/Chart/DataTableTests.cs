using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class SvgDataTableTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForLineChartWithDataTable()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SvgDataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\DataTable_sheet1_{ix++}.svg", svg);
                }
            }
        }
    }
}
