using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
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

                //var ix = 4;
                //var c = ws.Drawings[ix]; 
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\ChartForSvg{i}.svg", svg);
                }
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
                var ix = 0;
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

                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\5.3-SampleLines{ix}.svg", svg);

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\5.3-SampleLines{i}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateBlazorSample1()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BlazorSample1.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];

                //var ix = 1;
                //var c = ws.Drawings[ix];
                //var svg = c.ToSvg();
                //SaveTextFileToWorkbook($"svg\\5.3-SampleLines{ix}.svg", svg);

                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var c = ws.Drawings[i];
                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\BlazorSample1{i}.svg", svg);
                }
            }   
        }
        [TestMethod]
        public async Task HtmlExportWithLineChart()
        {
            using (var package = new ExcelPackage())
            {
                var style = TableStyles.Dark3;
                var sheet = package.Workbook.Worksheets.Add("Html export sample 8");
                var csvFileInfo = new FileInfo(Path.Combine(_dataPath, $"currencies2011weekly.csv"));
                if (csvFileInfo.Exists == false) return;
                var format = new ExcelTextFormat
                {
                    Delimiter = ';',
                    Culture = CultureInfo.InvariantCulture,
                    DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
                };
                var tableRange = sheet.Cells["A15"].LoadFromText(csvFileInfo, format, style, true);

                sheet.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[tableRange.Start.Row, 1, tableRange.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                sheet.Cells[tableRange.Start.Row, 2, tableRange.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";
                tableRange.AutoFitColumns();

                var table = sheet.Tables.GetFromRange(tableRange);
                table.ShowFirstColumn = true;
                var chart = sheet.Drawings.AddLineChart("LineChart1", eLineChartType.Line);

                var serie1 = chart.Series.Add(tableRange.TakeColumnsBetween(1,1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie1.HeaderAddress = sheet.Cells["B15"];

                var serie2 = chart.Series.Add(tableRange.TakeColumnsBetween(2, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie2.HeaderAddress = sheet.Cells["C15"];

                var serie3 = chart.Series.Add(tableRange.TakeColumnsBetween(3, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie3.HeaderAddress = sheet.Cells["D15"];

                chart.SetPosition(0, 0);
                chart.To.Row = 14;
                chart.To.Column = 10;
                chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle7);

                var exporter = sheet.Cells.CreateHtmlExporter();
                var settings = exporter.Settings;
                settings.Drawings.Include = eDrawingInclude.Include;
                settings.Culture = CultureInfo.InvariantCulture;
                settings.SetRowHeight = true;
                settings.SetColumnWidth = true;
                settings.TableId = "currency-table";
                settings.AdditionalTableClassNames.Add("table");
                settings.AdditionalTableClassNames.Add("table-sm");
                settings.AdditionalTableClassNames.Add("table-borderless");
                settings.Drawings.Position = eDrawingPosition.Absolute;
                SaveWorkbook("HtmlExportWithLineChart.xlsx", package);
                // export css and html
                //var css = exporter.GetCssString();
                //var html = exporter.GetHtmlString();
                var html = await exporter.GetSinglePageAsync();
                 
                SaveSvg("HtmlExportWithLineChart.html", html);
            }

        }
        //2.4-CreateAFileSystemReport.xlsx
        //3.3-FxReportFromDatabase.xlsx
    }
}
