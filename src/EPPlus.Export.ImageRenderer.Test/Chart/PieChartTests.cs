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

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieChartSvg{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GeneratePieChartFirstAngle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartSvgAngle.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieChartSvgAngle{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForPieChart()
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
                    SaveTextFileToWorkbook($"svg\\PieChartForSvg2{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForPieChartManySlices()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartManySlices.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieChartManySlices{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForPieChartFewSlices()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartLargeSlicesFewSeries.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieChartLargeSlicesFewSeries{ix++}.svg", svg);
                }
            }
        }


        [TestMethod]
        public void ReadAndGenerateExcelPieChartExplosion95()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieExplosion95Percent.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieExplosion95Percent_{ix++}.svg", svg);
                }
            }
        }


        [TestMethod]
        public void ReadAndGenerateExcelPieChartExplosionSvgs()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieExplosion30.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\PieExplosion30_{ix++}.svg", svg);
                }
            }
        }

    }
}
