using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
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
                for (int i = ix; i < ws.Drawings.Count; i++)
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

                var serie1 = chart.Series.Add(tableRange.TakeColumnsBetween(1, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie1.HeaderAddress = sheet.Cells["B15"];

                var serie2 = chart.Series.Add(tableRange.TakeColumnsBetween(2, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie2.HeaderAddress = sheet.Cells["C15"];

                var serie3 = chart.Series.Add(tableRange.TakeColumnsBetween(3, 1).SkipRows(1), tableRange.TakeColumns(1).SkipRows(1));
                serie3.HeaderAddress = sheet.Cells["D15"];

                chart.SetPosition(0, 0);
                chart.To.Row = 14;
                chart.To.Column = 10;
                chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle5);

                var textBox = sheet.Drawings.AddShape("InfoBox", eShapeStyle.RoundRect);
                textBox.RichText.Add("This is a line chart with data from the table below. The chart is exported as SVG when exporting to HTML. Sizes and positions in this export are absolute.");
                textBox.SetPosition(2, 0, 12, 0);
                textBox.SetSize(300, 200);

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
                //SaveWorkbook("HtmlExportWithLineChart.xlsx", package);
                // export css and html
                //var html = exporter.GetHtmlString();
                //var css = exporter.GetCssString();
                var html = await exporter.GetSinglePageAsync();

                SaveSvg("HtmlExportWithLineChart.html", html);
            }

        }
        [TestMethod]
        public async Task HtmlExportWithColumnGradient()
        {
            using var package = CreateWorkbook(eChartType.ColumnClustered, ePresetChartStyleMultiSeries.ColumnChartStyle9);
            var ws = package.Workbook.Worksheets[0];
            var svg = ws.Drawings[0].ToSvg();
            SaveSvg("ColumnGradient.svg", svg);
        }
        [TestMethod]
        public async Task HtmlExportWithPieWithDatalabels()
        {
            using var package = CreateWorkbook(eChartType.PieExploded, ePresetChartStyleMultiSeries.PieChartStyle7);
            var ws = package.Workbook.Worksheets[0];
            var svg = ws.Drawings[0].ToSvg();

            var of = GetOutputFile("", "PieChartHtml_dlbls.xlsx");
            package.SaveAs(of);

            var pChart = ws.Drawings[0].As.Chart.PieChart;
            pChart.DataLabel.ShowLegendKey = true;

            var svgWithLegendKey = ws.Drawings[0].ToSvg();

            var of2 = GetOutputFile("", "PieChartHtml_dlbls_legendKey.xlsx");
            package.SaveAs(of2);

            SaveSvg("PieWithDataLabels.svg", svg);
            SaveSvg("PieWithDataLabels_WithLegendKey.svg", svgWithLegendKey);
        }

        [TestMethod]
        public async Task SavingWorkbookPieChartDatalabels()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("ws1");
            ws.Cells["A1"].Value = 5;
            ws.Cells["A2"].Value = 10;
            ws.Cells["A3"].Value = 15;

            var myPie = ws.Drawings.AddPieChart("myPie", ePieChartType.Pie);
            myPie.Series.Add(ws.Cells["A1:A3"].TakeSingleColumn(0));

            myPie.DataLabel.ShowPercent = true;

            var of = GetOutputFile("", "PieChartHtml_dlbls_simple.xlsx");
            p.SaveAs(of);
        }

        public class RegionalSales
        {
            public string Region { get; set; }
            public int SoldUnits { get; set; }
            public double TotalSales { get; set; }
            public double Margin { get; set; }
        }

        private static List<RegionalSales> _salesData = new List<RegionalSales>()
        {
                new RegionalSales(){ Region = "North", SoldUnits=500, TotalSales=4800, Margin=0.200 },
                new RegionalSales(){ Region = "Central", SoldUnits=900, TotalSales=7330, Margin=0.333 },
                new RegionalSales(){ Region = "South", SoldUnits=400, TotalSales=3700, Margin=0.150 },
                new RegionalSales(){ Region = "East", SoldUnits=350, TotalSales=4400, Margin=0.102 },
                new RegionalSales(){ Region = "West", SoldUnits=700, TotalSales=6900, Margin=0.218 },
                new RegionalSales(){ Region = "Stockholm", SoldUnits=1200, TotalSales=8250, Margin=0.350 }
        };

        public static ExcelPackage CreateWorkbook(eChartType? chartType, ePresetChartStyleMultiSeries chartStyle)
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Html export with svg chart");

            var range = sheet.Cells["A1"].LoadFromCollection(_salesData, true, TableStyles.Dark3);
            sheet.Cells["B2:C7"].Style.Numberformat.Format = "#,##0";
            sheet.Cells["D2:D7"].Style.Numberformat.Format = "#,##0.00%";
            ExcelChart chart;
            switch (chartType)
            {
                case eChartType.LineMarkers:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.LineMarkers);
                    chart.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    break;
                case eChartType.ColumnClustered:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.ColumnClustered);
                    chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    break;
                case eChartType.BarClustered:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.BarClustered);
                    chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    break;
                case eChartType.PieExploded:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.PieExploded);
                    chart.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    var pieChart = chart as ExcelPieChart;
                    chart.StyleManager.SetChartStyle(chartStyle);
                    pieChart.DataLabel.ShowPercent = true;
                    pieChart.DataLabel.Border.Fill.Style = eFillStyle.SolidFill;
                    pieChart.DataLabel.Border.Fill.Color = Color.Black;
                    pieChart.DataLabel.Fill.Style = eFillStyle.SolidFill;
                    pieChart.DataLabel.Fill.Color = Color.LightCoral;
                    chart.SetPosition(2, 0, 5, 0);
                    chart.SetSize(1100, 300);
                    return package;
                    break;
                default:
                    chart = sheet.Drawings.AddChart("RegionalSalesChart", eChartType.ColumnClustered);
                    chart.Series.Add(sheet.Cells["B2:B7"], sheet.Cells["A2:A7"]);
                    chart.Series.Add(sheet.Cells["C2:C7"], sheet.Cells["A2:A7"]);
                    var lineChartType = chart.PlotArea.ChartTypes.Add(eChartType.Line);

                    lineChartType.UseSecondaryAxis = true;
                    lineChartType.Series.Add(sheet.Cells["D2:D7"], sheet.Cells["A2:A7"]);
                    break;

            }

            chart.StyleManager.SetChartStyle(chartStyle);
            chart.SetPosition(2, 0, 5, 0);
            chart.SetSize(1100, 400);

            return package;
        }

    }
}
