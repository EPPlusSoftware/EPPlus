using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_ExtLstTests : TestBase
    {
        [TestMethod]
        public void PriorityTestExtLst()
        {
            using (var pck = OpenPackage("CFPriorityTestExtLst.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("priorityTest");

                sheet.Cells["A1:A7"].Formula = "=Row()";
                sheet.Cells["A1:A7"].AutoFitColumns();
                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Green);

                var cfMiddlePriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A6"), Color.Yellow);

                var cfLowestPriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A7"), Color.Red);

                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A5"));

                cfHighestPriority.Formula = "0";
                cfHighestPriority.Style.Fill.BackgroundColor.Color = Color.Orange;

                var cfMiddlePriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A6"));

                cfMiddlePriority.Formula = "0";
                cfMiddlePriority.Style.Fill.BackgroundColor.Color = Color.Silver;

                var cfLowestPriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A7"));
                cfLowestPriority.Style.Fill.BackgroundColor.Color = Color.Yellow;
                cfLowestPriority.Formula = "0";

                SaveAndCleanup(pck);

                var readPck = OpenPackage("CFPriorityTestExtLst.xlsx");

                var cfs = readPck.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(4, cfs[0].Priority);
                Assert.AreEqual(5, cfs[1].Priority);
                Assert.AreEqual(6, cfs[2].Priority);
                Assert.AreEqual(1, cfs[3].Priority);
                Assert.AreEqual(2, cfs[4].Priority);
                Assert.AreEqual(3, cfs[5].Priority);
            }
        }

        [TestMethod]
        public void WriteReadEqualExt()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Equal");
                var ws2 = p.Workbook.Worksheets.Add("EqualExt");

                var cf = ws.Cells["A1"].ConditionalFormatting.AddEqual();
                cf.Formula = "EqualExt!A1";

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.Equal;
                    Assert.AreEqual("EqualExt!A1", cf.Formula);
                }
            }
        }

        [TestMethod]
        public void ExtLstFormulaValidations()
        {
            using (var pck = OpenPackage("ExtLstFormulas.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                refSheet.Cells["B5"].Value = 5;

                sheet.Cells["B1:B5"].Value = 5;
                sheet.Cells["B3"].Value = 2;

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Fill.BackgroundColor.Color = Color.Blue;
                equal.Style.Font.Italic = true;

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ExtLstWithDxf()
        {
            using (var pck = OpenPackage("ExtLstFormulasDxf.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Fill.BackgroundColor.Color = Color.Blue;
                equal.Style.Font.Italic = true;
                equal.Style.Font.Bold = false;

                var equal2 = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("C1:C5"));
                equal2.Formula = "formulasReference!$B$1";
                equal2.Style.Fill.Style = OfficeOpenXml.Style.eDxfFillStyle.GradientFill;
                var c1 = equal2.Style.Fill.Gradient.Colors.Add(0);
                var c2 = equal2.Style.Fill.Gradient.Colors.Add(100);

                equal2.Style.Fill.Gradient.Degree = 90;

                c1.Color.SetColor(Color.LightGreen);
                c2.Color.SetColor(Color.MediumPurple);

                SaveAndCleanup(pck);

                var readPck = OpenPackage("ExtLstFormulasDxf.xlsx", false);

                var format = readPck.Workbook.Worksheets[0].ConditionalFormatting[0];
            }
        }


        [TestMethod]
        public void ExtLstWithDxfBorderAndNumFmt()
        {
            using (var pck = OpenPackage("ExtLstBordersNumFmt.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                equal.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                equal.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dotted;
                equal.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dashed;
                equal.Style.NumberFormat.Format = "YYYY";

                var equal2 = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("C1:C5"));
                equal2.Formula = "formulasReference!$B$1";
                equal2.Style.Border.BorderAround();

                SaveAndCleanup(pck);

                var pck2 = OpenPackage("ExtLstBordersNumFmt.xlsx");

                var sheet2 = pck2.Workbook.Worksheets[0];

                var formatting = sheet2.ConditionalFormatting[0];

                Assert.AreEqual(OfficeOpenXml.Style.ExcelBorderStyle.Thick, formatting.Style.Border.Left.Style);
                Assert.AreEqual(OfficeOpenXml.Style.ExcelBorderStyle.Thin, formatting.Style.Border.Right.Style);
                Assert.AreEqual(OfficeOpenXml.Style.ExcelBorderStyle.Dotted, formatting.Style.Border.Top.Style);
                Assert.AreEqual(OfficeOpenXml.Style.ExcelBorderStyle.Dashed, formatting.Style.Border.Bottom.Style);
                Assert.AreEqual("YYYY", formatting.Style.NumberFormat.Format);

                var formatting2 = sheet2.ConditionalFormatting[1];

                Assert.AreEqual("formulasReference!$B$1", formatting2.Formula);
                Assert.AreEqual(OfficeOpenXml.Style.ExcelBorderStyle.Thin, formatting2.Style.Border.Right.Style);
            }
        }

        [TestMethod]
        public void EnsureExtLstDXFBorderColorsReadWrite()
        {
            using (var pck = OpenPackage("ExtLstBordersDXFColor.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                equal.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                equal.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dotted;
                equal.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dashed;

                equal.Style.Border.Left.Color.Color = Color.Coral;
                equal.Style.Border.Top.Color.Theme = OfficeOpenXml.Drawing.eThemeSchemeColor.Accent3;
                equal.Style.Border.Bottom.Color.Auto = true;

                SaveAndCleanup(pck);

                var readPackage = OpenPackage("ExtLstBordersDXFColor.xlsx");

                var readSheet = readPackage.Workbook.Worksheets[0];
                var formatting = readSheet.ConditionalFormatting[0];

                Assert.AreEqual(formatting.Style.Border.Left.Color.Color, Color.FromArgb(255, Color.Coral));
                Assert.AreEqual(formatting.Style.Border.Right.Color.HasValue, false);
                Assert.AreEqual(formatting.Style.Border.Top.Color.Theme, eThemeSchemeColor.Accent3);
                Assert.AreEqual(formatting.Style.Border.Bottom.Color.Auto, true);

                SaveAndCleanup(readPackage);
            }
        }

        [TestMethod]
        public void EnsureExtLstDXFBorderColorsThemeReadWrite()
        {
            using (var pck = OpenPackage("ExtLstBordersDXFTheme.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";

                sheet.Workbook.ThemeManager.CreateDefaultTheme();

                equal.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, eThemeSchemeColor.Accent5);

                SaveAndCleanup(pck);

                var readPck = OpenPackage("ExtLstBordersDXFTheme.xlsx");

                var readSheet = readPck.Workbook.Worksheets[0];
                var formatting = readSheet.ConditionalFormatting[0];

                Assert.AreEqual(eThemeSchemeColor.Accent5, formatting.Style.Border.Left.Color.Theme);
                Assert.AreEqual(eThemeSchemeColor.Accent5, formatting.Style.Border.Right.Color.Theme);
                Assert.AreEqual(eThemeSchemeColor.Accent5, formatting.Style.Border.Top.Color.Theme);
                Assert.AreEqual(eThemeSchemeColor.Accent5, formatting.Style.Border.Bottom.Color.Theme);
            }
        }

        [TestMethod]
        public void ConditionalFormattingOnSameAddressExtWriteRead()
        {
            using (var pck = OpenPackage("CF_SameAddressExt.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";

                var rule2 = sheet.ConditionalFormatting.AddBetween(new ExcelAddress("B1:B5"));

                rule2.Formula = "formulasReference!$B$5";
                rule2.Formula2 = "formulasReference!$B$6";

                SaveAndCleanup(pck);

                //Can it be read
                var readPackage = OpenPackage("CF_SameAddressExt.xlsx");
                var sheet2 = readPackage.Workbook.Worksheets[0];
                Assert.AreEqual(sheet2.ConditionalFormatting[0].Type, eExcelConditionalFormattingRuleType.Equal);
                Assert.AreEqual(sheet2.ConditionalFormatting[1].Type, eExcelConditionalFormattingRuleType.Between);
            }
        }


        [TestMethod]
        public void ConditionalFormattingMultipleKindsOnSameAddressReadWrite()
        {
            using (var pck = OpenPackage("CF_SameAddressExtManyTypes.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var extSheet = pck.Workbook.Worksheets.Add("formulasRef");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasRef!$A$1";

                var rule2 = sheet.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("B1:B5"), eExcelconditionalFormatting3IconsSetType.Stars);

                var rule3 = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("B1:B5"), Color.BlueViolet);

                SaveAndCleanup(pck);

                var readPackage = OpenPackage("CF_SameAddressExtManyTypes.xlsx");

                var formats = readPackage.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(eExcelConditionalFormattingRuleType.Equal, formats[0].Type);
                Assert.AreEqual(eExcelConditionalFormattingRuleType.ThreeIconSet, formats[1].Type);
                Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, formats[2].Type);
            }
        }

        [TestMethod]
        public void CF_MinMaxColourScale()
        {
            using (var pck = OpenPackage("CF_ColourScaleInverseMinMax.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var sheet2 = pck.Workbook.Worksheets.Add("formulasExt");

                var formatting = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A5"));
                formatting.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
                formatting.LowValue.Type = eExcelConditionalFormattingValueObjectType.Max;
                formatting.HighValue.Formula = "formulasExt!A1";

                SaveAndCleanup(pck);

                var readPck = OpenPackage("CF_ColourScaleInverseMinMax.xlsx");
                var test = readPck.Workbook.Worksheets[0].ConditionalFormatting[0];

                Assert.AreEqual("formulasExt!A1", test.As.ThreeColorScale.HighValue.Formula);
            }
        }

        [TestMethod]
        public void ReadWriteAllIExcelConditionalFormattingWithText()
        {
            using (var pck = OpenPackage("CF_TExt.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var sheet2 = pck.Workbook.Worksheets.Add("formulasRef");

                var text = "\"IF(\"Yes\"=\"Yes\",\"Hi\",\"Bye\")\"";
                var formula = "IF(\"Yes\"=\"Yes\",\"Hi\",\"Bye\")";

                var formattingNot = sheet.ConditionalFormatting.AddNotContainsText(new ExcelAddress("A1"));
                formattingNot.Text = text;
                var extFormattingNot = sheet.ConditionalFormatting.AddNotContainsText(new ExcelAddress("B1:B5"));
                extFormattingNot.Formula = formula;

                var formattingContains = sheet.ConditionalFormatting.AddContainsText(new ExcelAddress("A1"));
                formattingContains.Text = text;
                var extFormattingContains = sheet.ConditionalFormatting.AddContainsText(new ExcelAddress("B1:B5"));
                extFormattingContains.Formula = formula;

                var formattingEnds = sheet.ConditionalFormatting.AddEndsWith(new ExcelAddress("A1"));
                formattingEnds.Text = text;
                var extFormattingEnds = sheet.ConditionalFormatting.AddEndsWith(new ExcelAddress("B1:B5"));
                extFormattingEnds.Formula = formula;

                var formattingBegins = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1"));
                formattingBegins.Text = text;
                var extFormattingBegins = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("B1:B5"));
                extFormattingBegins.Formula = formula;

                SaveAndCleanup(pck);

                var readPck = OpenPackage("CF_TExt.xlsx");

                var count = readPck.Workbook.Worksheets[0].ConditionalFormatting.Count;

                //Note that extLst items are read in after all "normal" items into the conditionalFormattingList.
                //So we read in extLst items starting from index count - 4 as we have 4 "normal" items.
                var textTestNot = readPck.Workbook.Worksheets[0].ConditionalFormatting[0];
                var extTestNot = readPck.Workbook.Worksheets[0].ConditionalFormatting[count - 4];

                Assert.AreEqual(text, textTestNot.As.NotContainsText.Text);
                Assert.AreEqual(formula, extTestNot.As.NotContainsText.Formula);

                var textTestContains = readPck.Workbook.Worksheets[0].ConditionalFormatting[1];
                var extTestContains = readPck.Workbook.Worksheets[0].ConditionalFormatting[count - 3];

                Assert.AreEqual(text, textTestContains.As.ContainsText.Text);
                Assert.AreEqual(formula, extTestContains.As.ContainsText.Formula);

                var textTestEnds = readPck.Workbook.Worksheets[0].ConditionalFormatting[2];
                var extTestEnds = readPck.Workbook.Worksheets[0].ConditionalFormatting[count - 2];

                Assert.AreEqual(text, textTestEnds.As.EndsWith.Text);
                Assert.AreEqual(formula, extTestEnds.As.EndsWith.Formula);

                var textTestBegins = readPck.Workbook.Worksheets[0].ConditionalFormatting[3];
                var extTestBegins = readPck.Workbook.Worksheets[0].ConditionalFormatting[count - 1];

                Assert.AreEqual(text, textTestBegins.As.BeginsWith.Text);
                Assert.AreEqual(formula, extTestBegins.As.BeginsWith.Formula);
            }
        }

        [TestMethod]
        public void CF_ColourScaleExt()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("colourScale");
                var extSheet = pck.Workbook.Worksheets.Add("extSheet");

                var colorScale = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A20"));

                for (int i = 1; i < 21; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                colorScale.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Num;

                colorScale.MiddleValue.Formula = "$B$2";

                colorScale.LowValue.Formula = "IF($B$5 < extSheet!A1, 5, 10)";

                colorScale.HighValue.Formula = "B6";

                //colorScale.LowValue.Color = Color.AliceBlue;
                colorScale.LowValue.ColorSettings.SetColor(eThemeSchemeColor.Accent3);
                colorScale.LowValue.ColorSettings.Tint = 0.5f;

                colorScale.MiddleValue.ColorSettings.Index = 4;
                colorScale.MiddleValue.ColorSettings.Tint = 1.0f;

                colorScale.HighValue.ColorSettings.Auto = true;

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var readPackage = new ExcelPackage(stream);

                var scale = readPackage.Workbook.Worksheets[0].ConditionalFormatting[0];

                var threeCol = scale.As.ThreeColorScale;

                Assert.AreEqual(scale.As.ThreeColorScale.MiddleValue.Formula, "$B$2");
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.Formula, "IF($B$5 < extSheet!A1, 5, 10)");
                Assert.AreEqual(scale.As.ThreeColorScale.HighValue.Formula, "B6");

                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Theme, eThemeSchemeColor.Accent3);
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Tint, 0.5f);

                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Index, 4);
                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Tint, 1.0f);

                Assert.AreEqual(threeCol.HighValue.ColorSettings.Auto, true);
            }
        }


        [TestMethod]
        public void CF_ExtAndLocalMostComplexTypes()
        {
            using (var pck = OpenPackage("complexTest.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extWorksheet");
                var ws2 = pck.Workbook.Worksheets.Add("formulaWs");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress("A1"), Color.Magenta);

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress("A1"), Color.LimeGreen);

                ws.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1:B3"), eExcelconditionalFormatting3IconsSetType.Signs);
                
                var cfIconSet = ws.ConditionalFormatting.AddFourIconSet(new ExcelAddress("C1:C10"), eExcelconditionalFormatting4IconsSetType.TrafficLights);

                cfIconSet.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheckSymbol;

                ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("C1:C10"), eExcelconditionalFormatting5IconsSetType.Boxes);

                ws.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("C1:C20"));

                var textContains = ws.ConditionalFormatting.AddTextContains(new ExcelAddress("A1:Z50"));

                textContains.Priority = 1;

                textContains.Text = "abc";

                textContains.Formula = "formulaWs!B3";

                textContains.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                textContains.Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent2;

                SaveAndCleanup(pck);

                var pckRead = OpenPackage("complexTest.xlsx");
                var readSheet = pckRead.Workbook.Worksheets[0];

                readSheet.ConditionalFormatting[2].Style.Fill.BackgroundColor.Color = Color.Red;

                Assert.AreEqual(readSheet.ConditionalFormatting[3].As.DataBar.FillColor.Color, Color.FromArgb(255, Color.LimeGreen));
                Assert.AreEqual(readSheet.ConditionalFormatting[6].Style.Fill.BackgroundColor.Theme, eThemeSchemeColor.Accent2);

                SaveAndCleanup(pckRead);
            }
        }


        [TestMethod]
        public void CF_ExtFormulaBetweenReadWrite()
        {
            using (var pck = OpenPackage("Cf_BetweenExt.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("AverageSheet");
                var extWS = pck.Workbook.Worksheets.Add("ExtSheet");

                var between = ws.ConditionalFormatting.AddBetween(new ExcelAddress("A1:A40"));
                var notBetween = ws.ConditionalFormatting.AddNotBetween(new ExcelAddress("B1:B40"));

                between.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.DarkVertical;
                between.Style.Fill.PatternColor.Color = Color.BlueViolet;

                between.Style.Fill.BackgroundColor.Color = Color.DarkRed;
                notBetween.Style.Fill.BackgroundColor.Color = Color.DarkGreen;

                between.Formula = "ExtSheet!$B$3";
                between.Formula2 = "ExtSheet!$B$5";

                notBetween.Formula = "ExtSheet!$B$3";
                notBetween.Formula2 = "ExtSheet!$B$5";

                extWS.Cells["B3"].Value = 5;
                extWS.Cells["B5"].Value = 27;

                ws.Cells["A1:B40"].Formula = "Row()+3";

                SaveAndCleanup(pck);

                var readPck = OpenPackage("Cf_BetweenExt.xlsx");
                var readSheet = readPck.Workbook.Worksheets[0];

                Assert.AreEqual("ExtSheet!$B$3", readSheet.ConditionalFormatting[0].Formula);
                Assert.AreEqual("ExtSheet!$B$5", readSheet.ConditionalFormatting[0].Formula2);
                Assert.AreEqual(Color.FromArgb(255, Color.BlueViolet), readSheet.ConditionalFormatting[0].Style.Fill.PatternColor.Color);
                Assert.AreEqual(OfficeOpenXml.Style.ExcelFillStyle.DarkVertical, readSheet.ConditionalFormatting[0].Style.Fill.PatternType);
                Assert.AreEqual(Color.FromArgb(255, Color.DarkRed), readSheet.ConditionalFormatting[0].Style.Fill.BackgroundColor.Color);

                Assert.AreEqual("ExtSheet!$B$3", readSheet.ConditionalFormatting[1].Formula);
                Assert.AreEqual("ExtSheet!$B$5", readSheet.ConditionalFormatting[1].Formula2);
                Assert.AreEqual(Color.FromArgb(255, Color.DarkGreen), readSheet.ConditionalFormatting[1].Style.Fill.BackgroundColor.Color);
            }
        }
    }
}
