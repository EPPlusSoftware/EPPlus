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
    public class LineChartToSvgTests : TestBase
    {
        [TestMethod]
        public void GenerateSvgForLineCharts_sheet1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                var ix = 4;
                var c = ws.Drawings[ix];
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                //var ix = 0;
                //foreach (ExcelChart c in ws.Drawings)
                //{
                //    var svg = renderer.RenderDrawingToSvg(c);
                //    SaveTextFileToWorkbook($"svg\\ChartForSvg{ix++}.svg", svg);
                //}
            }
        }

        [TestMethod]
        public void GenerateSvgForLineCharts_sheet2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\ChartForSvg{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForLineCharts_sheet3()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[2];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_Sheet3{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForLineCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("LineChartRenderTest.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\LineChartForSvg_Single{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\LineChartForSvg{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForCharts_SecondaryAxis()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg_SecondaryAxis.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 3;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_SecAxis{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForCharts_SecondaryAxis_sheet2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg_SecondaryAxis.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 0;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_Sheet2_SecAxis{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSimplestChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SimplestChart.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\SimplestChartTitle.svg", svg);
            }
        }


        [TestMethod]
        public void GenerateDataLabels()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("datalabelsSvg.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\datalabelsAttempt.svg", svg);
            }
        }



        [TestMethod]
        public void GenerateDataLabelsTrueMost()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("datalabelsSvgTrueMostWithFill.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\datalabelsSvgTrueMostWithFill.svg", svg);
            }
        }

        [TestMethod]
        public void GenerateDataLabelsTrueMostAndManualLayout()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("datalabelsSvgTrueMostWithFillANDManual.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\datalabelsSvgTrueMostWithFillAndManual.svg", svg);
            }
        }

        [TestMethod]
        public void GenerateDatalabelsLeaderLines()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("datalabelsSvgLeaderLinesAdjustedToBeSimilar.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\datalabelsSvgLeaderLines.svg", svg);
            }
        }


        [TestMethod]
        public void GenerateDatalabelsRightAlignedWithBg()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("datalabelsSvgRightAlignedWithBg.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\datalabelsSvgLeaderLinesBg.svg", svg);
            }
        }

        [TestMethod]
        public void GenerateSimpleLineChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("defChartLine3Points.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\defChartLine3Points.svg", svg);
            }
        }
    }
}
