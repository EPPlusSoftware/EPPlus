using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class BarChartTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForBarCharts1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BarChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\BarChartForSvg_sheet1_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForBarCharts2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BarChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\BarChartForSvg_sheet2_{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void DatalabelBarCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BarChartForSvgDatalabelsBasic.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var drawings = ws.Drawings;
                var ix = 1;

                for (int i = ix; i < drawings.Count; i++)
                {
                    var svg = drawings[i].ToSvg();
                    SaveTextFileToWorkbook($"svg\\BarChartDataLabels{ix++}.svg", svg);
                }
            }
        }


        [TestMethod]
        public void NegativeDatalabelBarCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("negativeDatalabels.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var drawings = ws.Drawings;
                var ix = 0;

                for (int i = ix; i < drawings.Count; i++)
                {
                    var svg = drawings[i].ToSvg();
                    SaveTextFileToWorkbook($"svg\\NegativeLabels{ix++}.svg", svg);
                }
            }
        }
    }
}
