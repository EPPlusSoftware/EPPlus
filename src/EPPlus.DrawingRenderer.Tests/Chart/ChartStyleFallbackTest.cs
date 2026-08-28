using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;
using tc = OfficeOpenXml.Utils.TypeConversion;
using System.Globalization;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;

namespace EPPlus.DrawingRenderer.Tests.Chart
{
    [TestClass]
    public class ChartStyleFallbackTest : TestBase
    {

        [AssemblyInitialize]
        public static async Task AssemblyInit(TestContext context)
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            CreatePathIfNotExists("StyleExamples\\");

        }

        private void CreateStyleExampleAndExportIt(string fileName, Func<List<string>, bool> assertIfTestSuccessful)
        {
            bool testSucceded = false;

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                List<string> outputSvgs = new List<string>();

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        var svg = c.ToSvg();
                        outputSvgs.Add(svg);
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }

                testSucceded = assertIfTestSuccessful(outputSvgs);

                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
            Assert.IsTrue(testSucceded);
        }


        [TestMethod]
        public void ReadExcelFile()
        {
            CreateStyleExampleAndExportIt("ExcelUnchangedEmptyChart", (List<string> outputSvgs) => 
            {
                //Create expected color
                var col = Color.FromArgb(217, 217, 217);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 -1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(1d, widthResult);

                return expectedStr == colorResult && 1d == widthResult;
            });
        }

        [TestMethod]
        public void ReadEmptyDefaultChartStyle()
        {
            CreateStyleExampleAndExportIt("emptyDefault", (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(217, 217, 217);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(13.3333d, widthResult, 0.003);

                return expectedStr == colorResult && 13.3333d == Math.Round(widthResult,4);
            });
        }


        [TestMethod]
        public void ReadExcelEditedRemovedStyles()
        {
            string fileName = "emptyManuallyRemovedLnStyles";

            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(137, 137, 137);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(13.3333d, widthResult, 0.003);

                return expectedStr == colorResult && 13.3333d == Math.Round(widthResult, 4);
            });
        }

        [TestMethod]
        public void EditedTheme()
        {
            string fileName = "ExcelThemeEdited";

            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(255, 255, 199, 199);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(13.3333d, widthResult, 0.003);

                return expectedStr == colorResult && 13.3333d == Math.Round(widthResult, 4);
            });
        }


        [TestMethod]
        public void ManualSystemText()
        {
            string fileName = "ExcelThemeManualSystemText";

            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(255, 0, 0, 0);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(13.3333d, widthResult, 0.003);

                return expectedStr == colorResult && 13.3333d == Math.Round(widthResult, 4);
            });
        }


        [TestMethod]
        public void ExcelThemeLnDeleted()
        {
            string fileName = "ExcelThemeLnDeleted";

            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.Transparent;
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 4).ToLower();

                Assert.AreEqual("none", colorResult);

                return "none" == colorResult;
            });
        }


        [TestMethod]
        public void PureExcelTheme()
        {
            string fileName = "PureExcelTheme";

            //Border and Fill for chartArea expected colors
            List<string> ExpectedColors = new List<string>() { "#ffcaca", "#bbd7f9" };

            //Test un-edited excel theme with custom colors set in excel
            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                foreach(var svg in outputSvgs)
                {
                    var svgSplitOnSpace = svg.Split(' ');

                    //Get the first stroke and extract the hexCode for the expected color
                    var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                    var borderResult = firstStroke.Substring(8, 7).ToLower();

                    //Get the first stroke and extract the hexCode for the expected color
                    var firstFill = svgSplitOnSpace.First(s => s.StartsWith("fill"));
                    var fillResult = firstFill.Substring(6, 7).ToLower();

                    Assert.AreEqual(ExpectedColors[0], borderResult);
                    Assert.AreEqual(ExpectedColors[1], fillResult);
                }

                return true;
            });
        }


        [TestMethod]
        public void ChartWithChartStyle()
        {
            string fileName = "ChartWithChartStyleMEdit";

            //Read test for if chart style is applied appropriately
            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(255, 217, 217, 217);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(1d, widthResult, 0.003);

                return expectedStr == colorResult && 1d == Math.Round(widthResult, 1);
            });
        }

        [TestMethod]
        public void ReadChartBorderThemeTint()
        {
            var fileName = "ChartBorderThemeTint";

            //Read test for if chart style is applied appropriately
            CreateStyleExampleAndExportIt(fileName, (List<string> outputSvgs) =>
            {
                //Create expected color
                var col = Color.FromArgb(255, 217, 217, 217);
                var expectedStr = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = outputSvgs[0].Split(' ');

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var colorResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedStr, colorResult);
                Assert.AreEqual(1d, widthResult, 0.003);

                return expectedStr == colorResult && 1d == Math.Round(widthResult, 1);
            });
            //using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            //{
            //    var ws = p.Workbook.Worksheets[0];
            //    var lChart = ws.Drawings[0].As.Chart.LineChart;

            //    lChart.StyleManager.Style.ChartArea.Border.Fill.SolidFill.Color.SetSchemeColor(OfficeOpenXml.Drawing.eSchemeColor.Accent1);

            //    //100 - input is what excel seems to apply
            //    //lChart.StyleManager.Style.ChartArea.BorderReference.Color.Transforms.AddTint(13);

            //    //Adding Less Tint makes the object Lighter. Which is the inverse of how excel does it.
            //    lChart.StyleManager.Style.ChartArea.Border.Fill.SolidFill.Color.Transforms.AddTint(60);
            //    lChart.StyleManager.Style.ChartArea.Border.Width = 10d;
            //    lChart.StyleManager.ApplyStyles();

            //    var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
            //    p.SaveAs(fi);
            //}
        }


        [TestMethod]
        public void Epp_Gen_DefaultLine()
        {
            using (var p = OpenPackage("StyleExamples\\epplusDefaultChart_Line.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Empty");

                ws.Cells["A1:A3"].Formula = "ROW()+COLUMN()";

                ws.Calculate();

                var emptyLines = ws.Drawings.AddLineChart("Chart", eLineChartType.Line);
                foreach (ExcelDrawing d in ws.Drawings)
                {
                    var svg = d.ToSvg();

                    SaveTextFileToWorkbook($"svg\\epplusDefaultChart_Line{ws.Name}_{d.Name}.svg", svg);

                    //Create expected color
                    var fill = Color.FromArgb(255, 255, 255, 255);
                    var expectedFill = ColorTranslator.ToHtml(fill).ToLower();
                    var col = Color.FromArgb(255, 137, 137, 137);
                    var expectedStroke = ColorTranslator.ToHtml(col).ToLower();

                    var svgSplitOnSpace = svg.Split(' ');

                    //Get the first stroke and extract the hexCode for the expected color
                    var firstFill = svgSplitOnSpace.First(s => s.StartsWith("fill"));
                    var fillResult = firstFill.Substring(6, 7).ToLower();

                    //Get the first stroke and extract the hexCode for the expected color
                    var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                    var strokeResult = firstStroke.Substring(8, 7).ToLower();

                    //Get the resulting width
                    var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                    var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                    var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                    //Assert
                    Assert.AreEqual(expectedFill, fillResult);
                    Assert.AreEqual(expectedStroke, strokeResult);
                    Assert.AreEqual(1d, widthResult, 0.003);
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void Epp_Gen_DefaultShape()
        {
            string fileName = "epplusShape";

            using (var p = OpenPackage($"StyleExamples\\{fileName}.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("ws1");

                var defaultRect = ws.Drawings.AddShape("Default", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);
                var gradientRect = ws.Drawings.AddShape("GradRect", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);

                var aThemeStyle = defaultRect.ThemeStyles.BorderReference;

                defaultRect.SetPosition(300, 1);
                gradientRect.SetPosition(300, 1000);

                var svgDefault = defaultRect.ToSvg();
                SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{defaultRect.Name}.svg", svgDefault);

                //Create expected color
                var fill = Color.FromArgb(255, 21, 96, 130);
                var expectedFill = ColorTranslator.ToHtml(fill).ToLower();
                var col = Color.FromArgb(255, 4, 36, 51);
                var expectedStroke = ColorTranslator.ToHtml(col).ToLower();

                var svgSplitOnSpace = svgDefault.Split(' ');

                //Get the first fill and extract the hexCode for the expected color
                var firstFill = svgSplitOnSpace.First(s => s.StartsWith("fill"));
                var fillResult = firstFill.Substring(6, 7).ToLower();

                //Get the first stroke and extract the hexCode for the expected color
                var firstStroke = svgSplitOnSpace.First(s => s.StartsWith("stroke"));
                var strokeResult = firstStroke.Substring(8, 7).ToLower();

                //Get the resulting width
                var strokeWidth = svgSplitOnSpace.First(s => s.StartsWith("stroke-width"));
                var widthStr = strokeWidth.Substring(14, strokeWidth.Length - 14 - 1).ToLower();
                var widthResult = double.Parse(widthStr, CultureInfo.InvariantCulture);

                //Assert
                Assert.AreEqual(expectedFill, fillResult);
                Assert.AreEqual(expectedStroke, strokeResult);
                Assert.AreEqual(1d, widthResult, 0.003);

                gradientRect.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.GradientFill;

                var svgGradient = gradientRect.ToSvg();
                SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{gradientRect.Name}.svg", svgGradient);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void EpplusGeneratedChart()
        {
            using (var p = OpenPackage("StyleExamples\\epplusDefaultTest.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("EpplusGeneratedChart");

                //ws.Workbook.ThemeManager.GetOrCreateTheme();

                //ws.Workbook.ThemeManager.GetOrCreateTheme().FormatScheme.BackgroundFillStyle[0] = true;
                ws.Cells["A1:A3"].Formula = "ROW()+COLUMN()";

                ws.Calculate();

                var emptyLines = ws.Drawings.AddLineChart("EmptyLineChart", eLineChartType.Line);
                var generatedBar = ws.Drawings.AddBarChart("EpplusBarChart", eBarChartType.ColumnClustered);

                generatedBar.SetPosition(1, 1000);

                var defaultRect = ws.Drawings.AddShape("MyDefaultShape", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);
                var gradientRect = ws.Drawings.AddShape("GradRect", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);

                defaultRect.SetPosition(300, 1);
                gradientRect.SetPosition(300, 1000);

                defaultRect.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                gradientRect.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.GradientFill;
                generatedBar.Series.Add(ws.Cells["A1:A3"]);

                List<string> outputSvgs = new List<string>();

                foreach (ExcelDrawing d in ws.Drawings)
                {
                    var svg = d.ToSvg();
                    outputSvgs.Add(svg);
                    SaveTextFileToWorkbook($"svg\\epplusDefault{ws.Name}_{d.Name}.svg", svg);
                }
                //GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void GenerateSimpleChart()
        {
            string fileName = "EpplusSimpleChart";

            using (var p = OpenPackage($"{fileName}.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("s1");
                ws.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnClustered);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ReadObjectDefaults()
        {
            string fileName = "ObjectDefaultsChanged";

            using (var p = OpenTemplatePackage($"{fileName}.xlsx"))
            {
                var theme = p.Workbook.ThemeManager.GetOrCreateTheme();
                var shapeStyle = theme.ObjectDefaults.ShapeDefinition.Style;
                var fillRef = shapeStyle.FillReference;

                var lnRef = shapeStyle.BorderReference;

                Assert.AreEqual(eSchemeColor.Accent2, lnRef.ShapeColor.SchemeColor.Color);
                Assert.IsTrue(lnRef.ShapeColor.Transforms.Count > 0);
                Assert.AreEqual(eColorTransformType.Shade, lnRef.ShapeColor.Transforms[0].Type);
                Assert.AreEqual(15, lnRef.ShapeColor.Transforms[0].Value);
                Assert.AreEqual(2, lnRef.Index);

                Assert.IsTrue(fillRef.HasColor);
                var schemeClr = fillRef.ShapeColor.SchemeColor;
                var col = schemeClr.Color;
                Assert.AreEqual(eSchemeColor.Accent2, col);
                Assert.AreEqual(1, fillRef.Index);

                var fontRef = theme.ObjectDefaults.ShapeDefinition.Style.FontReference;
                Assert.AreEqual(eSchemeColor.Light1, fontRef.Color.SchemeColor.Color);
                Assert.AreEqual(eThemeFontCollectionType.Minor, fontRef.Index);

                SaveAndCleanup(p);
            }
        }
    }
}
