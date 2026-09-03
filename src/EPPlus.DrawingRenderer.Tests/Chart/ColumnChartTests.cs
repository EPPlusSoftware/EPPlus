using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class ColumnChartTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForColumnCharts1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ColumnChartForSvg.xlsx"))
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
                    SaveTextFileToWorkbook($"svg\\ColumnChartForSvg_sheet1_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForColumnCharts2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ColumnChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\ColumnChartForSvg_sheet2_{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateColumnChartFromBlazorSample()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BlazorSample1-Column.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\BlazorSample_Column_sheet2_{ix++}.svg", svg);
                }
            }            
        }
        [TestMethod]
        public void GenerateBarChartFromBlazorSample()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BlazorSample1-BarChart.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\BlazorSample_Bar_sheet2_{ix++}.svg", svg);
                }
            }
        }

    }
}
