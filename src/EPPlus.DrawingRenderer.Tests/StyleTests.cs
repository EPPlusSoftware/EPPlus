using OfficeOpenXml;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class StyleTests : TestBase
    {

        [TestMethod]
        public void BaseThemeChartStyle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("baseThemeChartStyle.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;
                var svg = c.ToSvg();
                //var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\baseThemeChartStyle.svg", svg);
            }
        }

        [TestMethod]
        public void BaseThemeChartStyle2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("baseThemeChartStyle2.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;
                var svg = c.ToSvg();
                //var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\baseThemeChartStyle2.svg", svg);
            }
        }
        [TestMethod]
        public void ChangeSeriesTest()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("MyLineIonThemeExcel.xlsx"))
            {
                var wbStyles = p.Workbook.Styles;
                var simpleChart = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;

                //simpleChart.StyleManager.ApplyStyles();

                p.Workbook.Worksheets[0].Cells["A3"].Value = 15;
                p.Workbook.Worksheets[0].Cells["B3"].Value = 2;
                p.Workbook.Worksheets[0].Cells["C3"].Value = 25;

                var range = p.Workbook.Worksheets[0].Cells["A3:C3"];

                var series = simpleChart.Series.Add(range);
                series.Header = "MySeries";

                simpleChart.Title.Text = "Hello";

                //simpleChart.StyleManager.ApplyStyles();
                //var chartDefaultStyle = simpleChart.StyleManager.Style;

                //simpleChart.StyleManager.Style.Title.FontReference.Color.SetPresetColor(Color.CornflowerBlue);

                ////Highest order of styling if datapoint does not exist
                ////simpleChart.StyleManager.Style.DeleteAllNode("cs:dataPointLine");
                //simpleChart.StyleManager.Style.DataPointLine.BorderReference.Color.SetPresetColor(Color.Green);

                //simpleChart.StyleManager.ApplyStyles();

                var svg = simpleChart.ToSvg();
                SaveTextFileToWorkbook($"svg\\ChangeSeriesStyleTest.svg", svg);
                p.SaveAs(GetOutputFile("", "ChangeSeriesStyleTest.xlsx"));
                //SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ExtractThemeStyleWorks()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("MyLineIonThemeExcel.xlsx"))
            {
                var wbStyles = p.Workbook.Styles;
                var simpleChart = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;

                //p.Workbook.Worksheets[0].Cells["A3"].Value = 15;
                //p.Workbook.Worksheets[0].Cells["B3"].Value = 2;
                //p.Workbook.Worksheets[0].Cells["C3"].Value = 25;

                //var range = p.Workbook.Worksheets[0].Cells["A3:C3"];

                //var series = simpleChart.Series.Add(range);
                //series.Header = "MySeries";

                simpleChart.Title.Text = "Hello";

                //p.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetPresetColor(Color.DarkGoldenrod);

                var chartDefaultStyle = simpleChart.StyleManager.Style;
                //simpleChart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.LineChartStyle2);
                ////var fillStyle = simpleChart.StyleManager.Style.DataPointLine.Border.Fill;

                //var svgFill = new SvgFill(fillStyle);
                //var fillTranslator = new SvgFillTranslator(svgFill);

                //var context = new TranslatorContext(new HtmlRangeExportSettings());

                //var declarations = fillTranslator.GenerateDeclarationList(context);

                //var styleClass = new CssRule("Style", 0);

                //context.SetTranslator(fillTranslator);
                //context.AddDeclarations(styleClass);

                //simpleChart.StyleManager.ApplyStyles();

                //chartDefaultStyle.DataPointLine.FillReference.Color
                //simpleChart.Title.Font.Color = Color.CornflowerBlue;
                simpleChart.StyleManager.Style.Title.FontReference.Color.SetPresetColor(Color.CornflowerBlue);

                //Highest order of styling if datapoint does not exist
                //simpleChart.StyleManager.Style.DeleteAllNode("cs:dataPointLine");
                simpleChart.StyleManager.Style.DataPointLine.BorderReference.Color.SetPresetColor(Color.Green);

                //
                //var currBorderFill = simpleChart.StyleManager.Style.DataPointLine.Border.Fill;
                //simpleChart.StyleManager.Style.DataPointLine.Border.Fill.DeleteNode(path);
                //simpleChart.StyleManager.Style.DataPointLine.Border.Fill.Color = Color.Red;
                //simpleChart.StyleManager.Style.DataPointLine.Fill.Color = Color.Green;

                //simpleChart.StyleManager.Style.DataPoint.Fill.Color = Color.Yellow;
                //simpleChart.StyleManager.Style.DataPoint.Border.Fill.Color = Color.Magenta;

                simpleChart.StyleManager.ApplyStyles();

                //simpleChart.Title.TextBody.Paragraphs[0].TextRuns[0].Fill.Color = Color.CornflowerBlue;

                var svg = simpleChart.ToSvg();
                SaveTextFileToWorkbook($"svg\\MyLineIonThemeExcel2.svg", svg);

                //var serLine = chartDefaultStyle.SeriesLine;
                //var lineElement = serLine.Border.LineElement;
                //var serLineFillRef = serLine.Border.Fill;
                //var myFillRef = chartDefaultStyle.Title.FillReference;
                //var myFill = chartDefaultStyle.Title.Fill;
                //var myLine = chartDefaultStyle.Title.Border;

                SaveAndCleanup(p);
            }

        }

        [TestMethod]
        public void ExtractThemeStyleWorksDataLine()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("MyLineIonThemeExcel.xlsx"))
            {
                var wbStyles = p.Workbook.Styles;
                var simpleChart = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;

                simpleChart.Title.Text = "Hello";

                //var chartDefaultStyle = simpleChart.StyleManager.Style;
                //simpleChart.StyleManager.ApplyStyles();

                //simpleChart.Title.TextBody.Paragraphs[0].TextRuns[0].Fill.Color = Color.CornflowerBlue;

                //var svg = simpleChart.ToSvg();
                //SaveTextFileToWorkbook($"svg\\MyLineIonThemeExcelChangeOnlyTitle.svg", svg);

                var fontSize = simpleChart.Title.Font.Size;

                Assert.AreEqual(simpleChart.Title.Font.Size, 14d);
                var paragraphPropeties = simpleChart.Title.GetNode("c:txPr/a:p/a:pPr");
                var paragraphPropertiesRich = simpleChart.Title.GetNode("c:tx/c:rich/a:p/a:pPr");
                Assert.AreEqual(paragraphPropeties.InnerXml, paragraphPropertiesRich.InnerXml);

                SaveAndCleanup(p);
            }

        }

        [TestMethod]
        public void TextRunIsStyledButNotTitleFont()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            var chartName = "ChartLineDefaultDownUp.xlsx";
            using (var p = OpenTemplatePackage(chartName))
            {
                var ws = p.Workbook.Worksheets[0];
                var lChart = ws.Drawings[0].As.Chart.LineChart;

                //lChart.Title.Font.Color = Color.DeepSkyBlue;

                lChart.StyleManager.ApplyStyles();

                SaveAndCleanup(p);
            }

            //using(var p= OpenPackage(chartName,false))
            //{

            //}
        }

        ///// <summary>
        ///// Exports an <see cref="ExcelTable"/> to a html string
        ///// </summary>
        ///// <returns>A html table</returns>
        //public string GetCssString()
        //{
        //    using (var ms = EPPlusMemoryManager.GetStream())
        //    {
        //        RenderCss(ms);
        //        ms.Position = 0;
        //        using (var sr = new StreamReader(ms))
        //        {
        //            return sr.ReadToEnd();
        //        }
        //    }
        //}
        ///// <summary>
        ///// Exports the css part of the html export.
        ///// </summary>
        ///// <param name="stream">The stream to write the css to.</param>
        ///// <exception cref="IOException"></exception>
        //public void RenderCss(Stream stream)
        //{
        //    var trueWriter = new CssWriter(stream);
        //    var cssRules = CreateRuleCollection(_settings);

        //    trueWriter.WriteAndClearFlush(cssRules, Settings.Minify);
        //}
    }
}

