using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class ErrorbarsTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForErrorbars_Line_Sheet1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Errorbars.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Error_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Errorbar_sheet1_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForErrorbars_Column_Sheet2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Errorbars.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Error_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Errorbar_sheet2_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForErrorbars_Bar_Sheet3()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("Errorbars.xlsx"))
            {
                var ws = p.Workbook.Worksheets[2];
                //var ix = 4;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\Error_sheet1_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\Errorbar_sheet3_{ix++}.svg", svg);
                }
            }
        }
    }
}
