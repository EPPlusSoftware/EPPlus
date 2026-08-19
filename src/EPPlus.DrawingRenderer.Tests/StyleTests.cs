using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;
using System.Linq;
using tc = OfficeOpenXml.Utils.TypeConversion;

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

        [TestMethod]
        public void
        ExcelShape()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            var fileName = "VbaShadedShape.xlsx";
            using (var p = OpenTemplatePackage(fileName))
            {
                var myShape = p.Workbook.Worksheets[0].Drawings[0].As.Shape;
                var expectedFill = Color.FromArgb(255, 21, 96, 130);
                var returned = tc.ColorConverter.ApplyTintDrawing(expectedFill, -0.85);
                var returnedAlt = tc.ColorConverter.ApplyTintDrawing(expectedFill, 0.15);
                //var myBlend = tc.ColorConverter.ApplyBlend(expectedFill, Color.Black, 0.15d);

                var outPutClr = tc.ColorConverter.Brighten(expectedFill, -0.85f);

                //myShape.Fill.SolidFill.Color

                //Expected hex input 156082
                //Expected hex output = 042433

                //All new Expected input #DB1BC0
                //All new expected output 0.6 tint lighten #F1CCE8
                //var myColor = ColorTranslator.FromHtml("#00FF00");


                //tc.ColorConverter.GetThemeColor
                //var outCol2 = tc.ColorConverter.ApplyLumMod(myColor, 0.4588235294117647d);
                //Expected result 241 204 232

                var shapeColor = myShape.Fill.SolidFill.Color;
            }
        }

        [TestMethod]
        public void EpplusGeneratedShape()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            var fileName = "Style_Epp_Rect.xlsx";
            using (var p = OpenPackage(fileName, true))
            {
                var ws = p.Workbook.Worksheets.Add("MyWs");
                var drawing = ws.Drawings.AddShape("rectangle", eShapeStyle.Rect);
                var expectedFill = Color.FromArgb(255, 21, 96, 130);
                var expectedBorder = Color.FromArgb(255, 4, 36, 51);

                var testTheme = tc.ColorConverter.ApplyTint(expectedFill, Convert.ToDouble(-0.85d));

                var docExample = Color.FromArgb(255, 79, 129, 189);
                ExcelDrawingRgbColor.GetHslColor(docExample, out double h, out double s, out double l);

                var expectedH = Math.Round(213d / 360d, 2);
                var expectedS = 0.45d;
                var expectedL = 0.53d;

                Assert.AreEqual(expectedH, Math.Round(h,2)/360d,0.05);
                Assert.AreEqual(expectedS, Math.Round(s, 2));
                Assert.AreEqual(expectedL, Math.Round(l, 2));

                var tint = 0.6d;
                var newLum = l * tint + (1 - tint);

                var newCol = ExcelDrawingHslColor.GetRgb(h, s, newLum);

                Assert.AreEqual(149, Convert.ToDouble(newCol.R));
                Assert.AreEqual(179, Convert.ToDouble(newCol.G));
                Assert.AreEqual(215, Convert.ToDouble(newCol.B));

                Single mySingle = 0.6F;

                var testAltTint = tc.ColorConverter.ApplyTintDrawing(docExample, Convert.ToDouble(mySingle));
                var retCol2= tc.ColorConverter.ApplyLumMod(testAltTint);

                //ExcelDrawingRgbColor.GetHslColor(Color.FromArgb(255,192,80,77), out double h, out double s, out double l);

                ////Schemecolor 56 what happens if 15%?
                ////00003366
                ////var test = tc.ColorConverter.ApplyBlend(expectedFill, Color.Black, 0.85d);
                ////var lum = l * 0.15d;

                ////var rgb = ExcelDrawingHslColor.GetRgb(h, s, 11.4d);
                ////var rgbAlt = ExcelDrawingHslColor.GetRgb(h, s, lum);

                var test = tc.ColorConverter.ApplyTintDrawing(Color.FromArgb(255, 0, 255, 0), 0.5d);

                //var test3 = tc.ColorConverter.ApplyTintDrawing(Color.FromArgb(255, 0, 255, 0), -0.5d);

                //var test2 = tc.ColorConverter.ApplyBlend(Color.FromArgb(255, 0, 255, 0), Color.Black, 0.5d);

                //var test4 = tc.ColorConverter.AlternativeTint(Color.FromArgb(255, 0, 255, 0), -0.5d);

                //var test5 = tc.ColorConverter.Brighten(Color.FromArgb(255, 0, 255, 0), -0.85f);

                //var excelWhite = Color.FromArgb(255, 245, 222, 179);

                //var test6 = tc.ColorConverter.Brighten(excelWhite, -0.5f);

                //var index56 = Color.FromArgb(0, 0, 51, 102);

                //var borderColorExcel = Color.FromArgb(0, 4, 36, 51);
                //var Accent1ColorExcel = Color.FromArgb(0, 21, 96, 130);

                //var shaded = tc.ColorConverter.ApplyLumMod(Accent1ColorExcel, 0.85d, 0.15d);

                //var appliedTint = tc.ColorConverter.ApplyTint(Accent1ColorExcel, -0.85d);

                var themeFill = tc.ColorConverter.GetThemeColor(ws.Workbook.ThemeManager.GetOrCreateTheme(), drawing.ThemeStyles.FillReference.Color);
                var themeBorder = tc.ColorConverter.GetThemeColor(ws.Workbook.ThemeManager.GetOrCreateTheme(), drawing.ThemeStyles.BorderReference.Color);

                Assert.AreEqual(expectedFill.ToArgb(), themeFill.ToArgb());
                Assert.AreEqual(expectedBorder.ToArgb(), themeBorder.ToArgb());

                var svg = drawing.ToSvg();
                SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{drawing.Name}.svg", svg);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void EpplusGeneratedShapeWithTheme()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            var fileName = "Style_Theme_Epp_Rect.xlsx";
            using (var p = OpenPackage(fileName, true))
            {
                var ws = p.Workbook.Worksheets.Add("MyWsWithTheme");
                var myThemeFile = GetTemplateFile("StyleExamples\\ParalaxTheme.thmx");
                p.Workbook.ThemeManager.Load(myThemeFile);

                var drawing = ws.Drawings.AddShape("rectangle", eShapeStyle.Rect);

                //var svg = drawing.ToSvg();
                //SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{drawing.Name}.svg", svg);

                SaveAndCleanup(p);
            }
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

