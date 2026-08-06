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

                var ix = 6;
                var c = ws.Drawings[ix];
                var svg = c.ToSvg();
                SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                //for (int i = 0; i < ws.Drawings.Count; i++)
                //{
                //    var c = ws.Drawings[i];
                //    var svg = c.ToSvg();
                //    SaveTextFileToWorkbook($"svg\\ChartForSvg{i}.svg", svg);
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
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
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
                //var ix = 0;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\chartforsvg_sheet3_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
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
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\LineChartForSvg_Single{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\LineChartForSvg{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForLineCharts3()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("LineChartRenderTest.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\LineChartForSvg_Single{ix++}.svg", svg);
                var ix = 2;
                for(int i = ix; i< ws.Drawings.Count; i++)
                {
                    var svg = ws.Drawings[i].ToSvg();
                    SaveTextFileToWorkbook($"svg\\LineChartForSvg{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSvgForLineChartSecondaryAxis()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg_SecondaryAxis.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
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
                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_Sheet2_SecAxis{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_Sheet2_SecAxis{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForCharts_SecondaryAxis_sheet3()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg_SecondaryAxis.xlsx"))
            {
                var ws = p.Workbook.Worksheets[2];
                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet3_{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\ChartForSvg_Sheet3_SecAxis{ix++}.svg", svg);
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
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

                var svg = c.ToSvg();
                SaveTextFileToWorkbook($"svg\\defChartLine3Points.svg", svg);
            }
        }
        [TestMethod]
        public void GenerateSvgForLineCharts_AxisAlign_sheet1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("HorizontalAxisAlign.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                //var ix = 3;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\HorizontalAxisChartForSvg{ix++}.svg", svg);

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\HorizontalAxisChartForSvg{i}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateEPPlusLineCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("3.3-FxReportFromDatabase.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\FxLineChart{i}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateLineChartWithDropLine()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("5.3-ChartsAndThemes-IntegralTheme.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\5.3-SampleLines{i}.svg", svg);
                }
            }
        }
        //2.4-CreateAFileSystemReport.xlsx
        //3.3-FxReportFromDatabase.xlsx
    }
}
