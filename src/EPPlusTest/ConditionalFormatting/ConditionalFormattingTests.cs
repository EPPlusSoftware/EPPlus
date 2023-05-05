/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using System.Xml.Linq;

namespace EPPlusTest.ConditionalFormatting
{
    /// <summary>
    /// Test the Conditional Formatting feature
    /// </summary>
    [TestClass]
    public class ConditionalFormattingTests : TestBase
    {
        private static ExcelPackage _pck;

        [ClassInitialize()]
        public static void Init(TestContext testContext)
        {
            _pck = OpenPackage("ConditionalFormatting.xlsx", true);
        }
        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void CleanUp()
        {
            SaveAndCleanup(_pck);
        }

        /// <summary>
        /// 
        /// </summary>
        [TestMethod]
        public void TwoColorScale()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColorScale");
            var cf = ws.ConditionalFormatting.AddTwoColorScale(ws.Cells["A1:A5"]);
            cf.PivotTable = true;
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);
            ws.SetValue(4, 1, 4);
            ws.SetValue(5, 1, 5);
        }
        [TestMethod]
        public void Pivot()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pivot");
            var cf = ws.ConditionalFormatting.AddThreeColorScale(ws.Cells["A1:A5"]);
            cf.PivotTable = false;
        }

        /// <summary>
        /// 
        /// </summary>
        [TestMethod]
        public void TwoBackColor()
        {
            var ws = _pck.Workbook.Worksheets.Add("TwoBackColor");
            IExcelConditionalFormattingEqual condition1 = ws.ConditionalFormatting.AddEqual(ws.Cells["A1"]);
            condition1.StopIfTrue = true;
            condition1.Priority = 1;
            condition1.Formula = "TRUE";
            condition1.Style.Fill.BackgroundColor.Color = Color.Green;
            IExcelConditionalFormattingEqual condition2 = ws.ConditionalFormatting.AddEqual(ws.Cells["A2"]);
            condition2.StopIfTrue = true;
            condition2.Priority = 2;
            condition2.Formula = "FALSE";
            condition2.Style.Fill.BackgroundColor.Color = Color.Red;
        }
        [TestMethod]
        public void Databar()
        {
            var ws = _pck.Workbook.Worksheets.Add("Databar");
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);
            ws.SetValue(4, 1, 4);
            ws.SetValue(5, 1, 5);
        }

        [TestMethod]
        public void DatabarChangingAddressCorrectly()
        {
            var ws = _pck.Workbook.Worksheets.Add("DatabarAddressing");
            // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
            ws.HeaderFooter.AlignWithMargins = true;
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            cf.Address = new ExcelAddress("C3");

            Assert.AreEqual(cf.Address, "C3");
        }

        [TestMethod]
        public void IconSet()
        {
            var ws = _pck.Workbook.Worksheets.Add("IconSet");
            var cf = ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);

            var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
            cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon1.Formula = "0";
            cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon2.Formula = "1/3";
            cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon3.Formula = "2/3";
            ws.SetValue(1, 2, 1);
            ws.SetValue(2, 2, 2);
            ws.SetValue(3, 2, 3);
            ws.SetValue(4, 2, 4);

            var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"], eExcelconditionalFormatting5IconsSetType.Quarters);
            cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon1.Value = 1;
            cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon2.Value = 2;
            cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon3.Value = 3;
            cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon4.Value = 4;
            cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon5.Value = 5;
            cf5.ShowValue = false;
            cf5.Reverse = true;

            ws.SetValue(1, 3, 1);
            ws.SetValue(2, 3, 2);
            ws.SetValue(3, 3, 3);
            ws.SetValue(4, 3, 4);
            ws.SetValue(5, 3, 5);
        }
        [TestMethod]
        public void WriteReadEqual()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Equal");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddEqual();
                cf.Formula = "1";
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.Equal;
                    Assert.AreEqual("1", cf.Formula);
                }
            }
        }

        [TestMethod]
        public void WriteReadThreeIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights2);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights2, cf.IconSet);
                }
            }
        }
        [TestMethod]
        public void WriteReadFourIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FourIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.ArrowsGray);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FourIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting4IconsSetType.ArrowsGray, cf.IconSet);
                }
            }
        }
        [TestMethod]
        public void WriteReadFiveIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFiveIconSet(eExcelconditionalFormatting5IconsSetType.Arrows);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FiveIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting5IconsSetType.Arrows, cf.IconSet);
                }
            }
        }


        [TestMethod]
        public void WriteReadDataBar()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("DataBar");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddDatabar(Color.Red);

                p.Save();

                //SaveAndCleanup(p);

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.DataBar;
                    Assert.AreEqual(Color.Red.ToArgb(), cf.Color.ToArgb());
                }
            }
        }

        [TestMethod]
        public void WriteReadTwoColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TwoColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddTwoColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.TwoColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void WriteReadThreeColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ThreeColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                cf.MiddleValue.Value = 25;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void VerifyReadStyling()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddBetween(ws.Cells["A1:A3"]);
                cf.Formula = "1";
                cf.Formula2 = "2";

                string expectedFormat = "#,##0";
                cf.Style.Font.Bold = true;
                cf.Style.Font.Italic = true;
                cf.Style.Font.Color.SetColor(Color.Red);
                cf.Style.NumberFormat.Format = expectedFormat;

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.Between;
                    Assert.IsTrue(cf.Style.Font.Bold.Value);
                    Assert.IsTrue(cf.Style.Font.Italic.Value);
                    Assert.AreEqual(Color.Red.ToArgb(), cf.Style.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(expectedFormat, cf.Style.NumberFormat.Format);
                }
            }
        }
        [TestMethod]
        public void VerifyExpression()
        {
            using (var p = OpenPackage("cf.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(new ExcelAddress("$1:$1048576"));
                cf.Formula = "IsError(A1)";
                cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void TestInsertRowsIntoVeryLongRangeWithConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the whole of column A except row 1
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A2:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert some rows into the worksheet
                wks.InsertRow(5, 3);

                // Check that the conditional formatting rule still applies to the same range (since there's nowhere to extend it to)
                Assert.AreEqual(cfAddress, cf.Address.Address);
            }
        }
        [TestMethod]
        public void TestInsertRowsAboveVeryLongRangeWithConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the whole of column A except rows 1-10
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A11:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert 3 rows into the worksheet above the conditional formatting
                wks.InsertRow(5, 3);

                // Check that the conditional formatting rule starts lower down, but ends in the same place
                Assert.AreEqual("A14:A1048576", cf.Address.Address);
            }
        }

        [TestMethod]
        public void TestInsertRowsToPushConditionalFormattingOffSheet()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the last two rows of column A
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A1048575:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(1, wks.ConditionalFormatting.Count);
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert enough rows into the worksheet above the conditional formatting rule to push it off the sheet 
                wks.InsertRow(5, 10);

                // Check that the conditional formatting rule no longer exists
                Assert.AreEqual(0, wks.ConditionalFormatting.Count);
            }
        }

        [TestMethod]
        public void TestNewConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the last two rows of column A
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A1:A10";

                for (int i = 1; i < 11; i++)
                {
                    wks.Cells[i, 1].Value = i;
                }

                var cf = wks.ConditionalFormatting.AddGreaterThan(new ExcelAddress(cfAddress));
                cf.Formula = "5.5";
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                cf.Style.Font.Color.SetColor(Color.White);

                pck.SaveAs("C:\\epplusTest\\Workbooks\\conditionalTest.xlsx");
            }
        }

        [TestMethod]
        public void CustomIconsWriteRead()
        {

            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var threeIcon = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1"), eExcelconditionalFormatting3IconsSetType.Triangles);

                threeIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedFlag;
                threeIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                threeIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow;

                var fourIcon = wks.ConditionalFormatting.AddFourIconSet(new ExcelAddress("A2"), eExcelconditionalFormatting4IconsSetType.Rating);

                fourIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fourIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fourIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fourIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;

                var fiveIcon = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("B1"), eExcelconditionalFormatting5IconsSetType.Boxes);

                fiveIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fiveIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fiveIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fiveIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
                fiveIcon.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircle;

                var specialCase = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("B1"), eExcelconditionalFormatting5IconsSetType.Boxes);

                specialCase.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                specialCase.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar;
                specialCase.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars;
                specialCase.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                specialCase.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars;

                MemoryStream stream = new MemoryStream();
                pck.SaveAs(stream);

                ExcelPackage package2 = new ExcelPackage(stream);

                var threeIconRead = (ExcelConditionalFormattingThreeIconSet)package2.Workbook.Worksheets[0].ConditionalFormatting[0];

                Assert.AreEqual(threeIconRead.Icon1.CustomIcon, eExcelconditionalFormattingCustomIcon.RedFlag);

                package2.SaveAs("C:/epplusTest/Workbooks/FlagTest.xlsx");
            }
        }

        [TestMethod]
        public void EnsureCustomIconsReturnCorrectStrings()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var validation = (ExcelConditionalFormattingThreeIconSet)wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1"), eExcelconditionalFormatting3IconsSetType.Triangles);

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowSideArrow;

                Assert.AreEqual("3Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayUpArrow;

                Assert.AreEqual("3ArrowsGray", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowFlag;
                Assert.AreEqual("3Flags", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCircle;

                Assert.AreEqual("3TrafficLights1", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowTrafficLight;
                Assert.AreEqual("3TrafficLights2", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDiamond;
                Assert.AreEqual("3Signs", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol;
                Assert.AreEqual("3Symbols", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheck;
                Assert.AreEqual("3Symbols2", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.HalfGoldStar;
                Assert.AreEqual("3Stars", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow;
                Assert.AreEqual("4Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                Assert.AreEqual("4RedToBlack", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                Assert.AreEqual("4Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                Assert.AreEqual("4TrafficLights", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                Assert.AreEqual("5Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters;
                Assert.AreEqual("5Quarters", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.OneFilledBox;
                Assert.AreEqual("5Boxes", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                Assert.AreEqual("NoIcons", validation.Icon1.GetCustomIconStringValue());
            }
        }


        [TestMethod]
        public void BeginsWith_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.BeginsWith;

            BaseReadWriteTest("A1:A5", "BeginsWith", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddBeginsWith(address);
                });
        }

        [TestMethod]
        public void EndsWith_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.EndsWith;

            BaseReadWriteTest("A1:A5", "EndsWith", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddEndsWith(address);
                });
        }


        [TestMethod]
        public void Expression_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.Expression;

            BaseReadWriteTest("A1:A5", "Expression", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddExpression(address);
                });
        }


        [TestMethod]
        public void GreaterThanOrEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.GreaterThanOrEqual;

            BaseReadWriteTest("A1:A5", "GreaterThanOrEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddGreaterThanOrEqual(address);
                });
        }

        [TestMethod]
        public void LessThanOrEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.LessThanOrEqual;

            BaseReadWriteTest("A1:A5", "LessThanOrEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddLessThanOrEqual(address);
                });
        }

        [TestMethod]
        public void ContainsBlanks_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsBlanks;

            BaseReadWriteTest("A1:A5", "ContainsBlanks", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddContainsBlanks(address);
                });
        }

        [TestMethod]
        public void NotBetween_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotBetween;

            BaseReadWriteTest("A1:A5", "NotBetween", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddNotBetween(address);
                    cf.Formula = "1";
                    cf.Formula2 = "5";

                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void ContainsErrors_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsErrors;

            BaseReadWriteTest("A1:A5", "ContainsErrors", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddContainsErrors(address);
                });
        }


        [TestMethod]
        public void Equal_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.Equal;

            BaseReadWriteTest("A1:A5", "Equal", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddEqual(address);
                });
        }

        [TestMethod]
        public void NotEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotEqual;

            BaseReadWriteTest("A1:A5", "NotEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotEqual(address);
                });
        }

        [TestMethod]
        public void UniqueValues()
        {
            var type = eExcelConditionalFormattingRuleType.UniqueValues;

            BaseReadWriteTest("A1:A5", "UniqueValues", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddUniqueValues(address);
                });
        }

        [TestMethod]
        public void ContainsText_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsText;

            BaseReadWriteTest("A1:A5", "ContainsText", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddContainsText(address);
                    cf.ContainText = "a";
                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void NotContainsText_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotContainsText;

            BaseReadWriteTest("A1:A5", "NotContainsText", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddNotContainsText(address);
                    cf.ContainText = "a";
                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void NotContainsBlanks_WriteRead()
        {
            BaseReadWriteTest("A1:A5", "NotContainsBlanks", eExcelConditionalFormattingRuleType.NotContainsBlanks,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotContainsBlanks(address);
                });
        }

        [TestMethod]
        public void NotContainsErrors_WriteRead()
        {
            BaseReadWriteTest("A1:A5", "NotContainsErrors", eExcelConditionalFormattingRuleType.NotContainsErrors,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotContainsErrors(address);
                });
        }

        private void BaseReadWriteTest(string address, string wsName, eExcelConditionalFormattingRuleType type,
            Func<ExcelWorksheet, ExcelAddress, ExcelConditionalFormattingRule> MakeCF)
        {
            ExcelWorksheet origWS, ws;
            var package = GenerateWorkSheets(wsName, out origWS, out ws);

            var cf1 = MakeCF(origWS, new ExcelAddress(address));
            var cf2 = MakeCF(ws, new ExcelAddress(address));

            ApplyColorStyle(cf1, cf2);
            TestReadWrite(package, cf2, type);
        }

        private ExcelPackage GenerateWorkSheets(string name, out ExcelWorksheet sheet1, out ExcelWorksheet sheet2)
        {
            var package = new ExcelPackage();
            sheet1 = _pck.Workbook.Worksheets.Add(name);
            sheet2 = package.Workbook.Worksheets.Add(name);

            return package;
        }

        private void ApplyColorStyle(ExcelConditionalFormattingRule origCF, ExcelConditionalFormattingRule cf)
        {
            origCF.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            origCF.Style.Fill.BackgroundColor.Color = Color.Aquamarine;

            cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cf.Style.Fill.BackgroundColor.Color = Color.Aquamarine;
        }

        private void TestReadWrite(ExcelPackage package, ExcelConditionalFormattingRule cf, eExcelConditionalFormattingRuleType type)
        {
            var stream = new MemoryStream();
            package.SaveAs(stream);

            ExcelPackage package2 = new ExcelPackage(stream);

            var cf2 = package.Workbook.Worksheets[0].ConditionalFormatting[0];

            Assert.AreEqual(cf.Formula, cf2.Formula);
            Assert.AreEqual(cf.Formula2, cf2.Formula2);

            Assert.AreEqual(cf.Text, cf2.Text);
            Assert.AreEqual(cf2.Type, type);

            var stream2 = new MemoryStream();
            package2.SaveAs("C:\\epplusTest\\Workbooks\\cf.xlsx");
        }


        static string[] numbers = new string[]
        { "zero",
          "one",
          "two",
          "three",
          "four",
          "five",
          "six",
          "seven",
          "eight",
          "nine",
          "ten",
          "eleven"
        };

        [TestMethod]
        public void TestReadingConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                string date = "2023-03-";

                string lastMonth = "2023-02-";
                string thisMonth = "2023-03-";
                string nextMonth = "2023-04-";

                for (int i = 1; i < 11; i++)
                {
                    wks.Cells[i, 5].Value = i;
                    wks.Cells[i, 6].Value = i;
                    wks.Cells[i, 8].Value = i % 2;
                    wks.Cells[i, 10].Value = numbers[i];

                    wks.Cells[i, 12].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 12].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 13].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 13].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 14].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 14].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 15].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 15].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 16].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 16].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 16].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 16].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 17].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 17].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 18].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 18].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 19].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 19].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 19].Value = nextMonth + $"{i + 10}";

                    wks.Cells[i, 20].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 20].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 20].Value = nextMonth + $"{i + 10}";

                    wks.Cells[i, 21].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 21].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 21].Value = nextMonth + $"{i + 10}";

                    int counter = 0;
                    wks.Cells[i, 23].Value = i % 2 == 1 ? i : counter++ % 2;

                    wks.Cells[i, 25].Value = i;
                    wks.Cells[i + 10, 25].Value = i + 10;

                    wks.Cells[i, 26].Value = i;
                    wks.Cells[i + 10, 26].Value = i + 10;

                    wks.Cells[i, 27].Value = i;
                    wks.Cells[i + 10, 27].Value = i + 10;

                    wks.Cells[i, 28].Value = i;
                    wks.Cells[i + 10, 28].Value = i + 10;

                    wks.Cells[i, 38].Value = i;

                    wks.Cells[i, 40].Value = i;
                    wks.Cells[i, 41].Value = i;

                    wks.Cells[i, 43].Value = i;
                    wks.Cells[i, 44].Value = i;
                    wks.Cells[i, 45].Value = i;
                    wks.Cells[i, 46].Value = i;
                    wks.Cells[i, 47].Value = i;
                }

                for (int i = 0; i < 4; i++)
                {
                    wks.Cells[1, 30 + i].Value = 3;
                    wks.Cells[2, 30 + i].Value = 2;
                    wks.Cells[3, 30 + i].Value = 4;
                }

                for (int i = 0; i < 2; i++)
                {
                    wks.Cells[1, 35 + i].Value = -500;
                    wks.Cells[2, 35 + i].Value = -10;
                    wks.Cells[3, 35 + i].Value = -1;
                    wks.Cells[4, 35 + i].Value = 0;
                    wks.Cells[5, 35 + i].Value = 1;
                    wks.Cells[6, 35 + i].Value = 9;
                    wks.Cells[7, 35 + i].Value = 17;
                    wks.Cells[8, 35 + i].Value = 25;
                    wks.Cells[9, 35 + i].Value = 200;
                }

                var betweenFormatting = wks.ConditionalFormatting.AddBetween(new ExcelAddress(1, 5, 10, 5));
                betweenFormatting.Formula = "3";
                betweenFormatting.Formula2 = "8";

                betweenFormatting.Style.Fill.BackgroundColor.Color = Color.Red;
                betweenFormatting.Style.Font.Color.Color = Color.Orange;

                var lessFormatting = wks.ConditionalFormatting.AddLessThan(new ExcelAddress(1, 6, 10, 6));
                lessFormatting.Formula = "7";

                lessFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                lessFormatting.Style.Font.Color.Color = Color.Violet;

                var equalFormatting = wks.ConditionalFormatting.AddEqual(new ExcelAddress(1, 8, 10, 8));
                equalFormatting.Formula = "1";

                equalFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                equalFormatting.Style.Font.Color.Color = Color.Violet;

                var containsFormatting = wks.ConditionalFormatting.AddTextContains(new ExcelAddress(1, 10, 10, 10));
                containsFormatting.ContainText = "o";

                containsFormatting.Style.Fill.BackgroundColor.Color = Color.Blue;
                containsFormatting.Style.Font.Color.Color = Color.Yellow;

                var dateFormatting = wks.ConditionalFormatting.AddLast7Days(new ExcelAddress(1, 12, 10, 12));

                dateFormatting.Style.Fill.BackgroundColor.Color = Color.Red;
                dateFormatting.Style.Font.Color.Color = Color.Yellow;

                var yesterdayFormatting = wks.ConditionalFormatting.AddYesterday(new ExcelAddress(1, 13, 10, 13));

                //TODO: Fix Priority. It doesn't seem to apply correctly.

                yesterdayFormatting.Style.Fill.BackgroundColor.Color = Color.Gray;
                yesterdayFormatting.Style.Font.Color.Color = Color.Red;
                yesterdayFormatting.Priority = 1;

                var todayFormatting = wks.ConditionalFormatting.AddToday(new ExcelAddress(1, 14, 10, 14));

                todayFormatting.Style.Fill.BackgroundColor.Color = Color.Yellow;
                todayFormatting.Style.Font.Color.Color = Color.Green;
                yesterdayFormatting.Priority = 2;

                var tomorrow = wks.ConditionalFormatting.AddTomorrow(new ExcelAddress(1, 15, 10, 15));

                tomorrow.Style.Fill.BackgroundColor.Color = Color.Black;
                tomorrow.Style.Font.Color.Color = Color.Violet;

                var lastWeek = wks.ConditionalFormatting.AddLastWeek(new ExcelAddress(1, 16, 20, 16));

                lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                lastWeek.Style.Font.Color.Color = Color.Violet;

                var thisWeek = wks.ConditionalFormatting.AddThisWeek(new ExcelAddress(1, 17, 20, 17));

                thisWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                thisWeek.Style.Font.Color.Color = Color.Violet;

                var nextWeek = wks.ConditionalFormatting.AddNextWeek(new ExcelAddress(1, 18, 20, 18));

                nextWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                nextWeek.Style.Font.Color.Color = Color.Violet;

                var lastMonthCF = wks.ConditionalFormatting.AddLastMonth(new ExcelAddress(1, 19, 27, 19));

                lastMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                lastMonthCF.Style.Font.Color.Color = Color.Violet;

                var thisMonthCF = wks.ConditionalFormatting.AddThisMonth(new ExcelAddress(1, 20, 27, 20));

                thisMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                thisMonthCF.Style.Font.Color.Color = Color.Violet;

                var nextMonthCF = wks.ConditionalFormatting.AddNextMonth(new ExcelAddress(1, 21, 27, 21));

                nextMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                nextMonthCF.Style.Font.Color.Color = Color.Violet;

                var duplicateValues = wks.ConditionalFormatting.AddDuplicateValues(new ExcelAddress(1, 23, 10, 23));

                duplicateValues.Style.Fill.BackgroundColor.Color = Color.Blue;
                duplicateValues.Style.Font.Color.Color = Color.Yellow;


                var top11 = wks.ConditionalFormatting.AddTop(new ExcelAddress(1, 25, 20, 25));

                top11.Rank = 11;
                top11.Style.Fill.BackgroundColor.Color = Color.Black;
                top11.Style.Font.Color.Color = Color.Violet;

                var bot12 = wks.ConditionalFormatting.AddBottom(new ExcelAddress(1, 26, 20, 26));

                bot12.Rank = 12;
                bot12.Style.Fill.BackgroundColor.Color = Color.Black;
                bot12.Style.Font.Color.Color = Color.Violet;

                var top13Percent = wks.ConditionalFormatting.AddTopPercent(new ExcelAddress(1, 27, 20, 27));

                top13Percent.Rank = 13;
                top13Percent.Style.Fill.BackgroundColor.Color = Color.Black;
                top13Percent.Style.Font.Color.Color = Color.Violet;

                var bot14Percent = wks.ConditionalFormatting.AddBottomPercent(new ExcelAddress(1, 28, 20, 28));

                bot14Percent.Rank = 14;
                bot14Percent.Style.Fill.BackgroundColor.Color = Color.Black;
                bot14Percent.Style.Font.Color.Color = Color.Violet;

                var aboveAverage = wks.ConditionalFormatting.AddAboveAverage(new ExcelAddress(1, 30, 20, 30));

                aboveAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveAverage.Style.Font.Color.Color = Color.Violet;

                var aboveOrEqualAverage = wks.ConditionalFormatting.AddAboveOrEqualAverage(new ExcelAddress(1, 31, 20, 31));

                aboveOrEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveOrEqualAverage.Style.Font.Color.Color = Color.Violet;

                var belowAverage = wks.ConditionalFormatting.AddBelowAverage(new ExcelAddress(1, 32, 20, 32));

                belowAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                belowAverage.Style.Font.Color.Color = Color.Violet;

                var belowEqualAverage = wks.ConditionalFormatting.AddBelowOrEqualAverage(new ExcelAddress(1, 33, 20, 33));

                belowEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                belowEqualAverage.Style.Font.Color.Color = Color.Violet;

                var aboveStdDev = wks.ConditionalFormatting.AddAboveStdDev(new ExcelAddress(1, 35, 10, 35));

                aboveStdDev.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveStdDev.Style.Font.Color.Color = Color.Violet;

                aboveStdDev.StdDev = 1;

                var belowStdDev = wks.ConditionalFormatting.AddBelowStdDev(new ExcelAddress(1, 36, 10, 36));

                belowStdDev.Style.Fill.BackgroundColor.Color = Color.Black;
                belowStdDev.Style.Font.Color.Color = Color.Violet;

                belowStdDev.StdDev = 2;

                var databar = wks.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 38, 10, 38), Color.AliceBlue);
                databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                databar.LowValue.Value = 0;
                databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                databar.HighValue.Value = 50;

                var twoColor = wks.ConditionalFormatting.AddTwoColorScale(new ExcelAddress(1, 40, 10, 40));

                twoColor.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                twoColor.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;

                twoColor.LowValue.Value = 5;
                twoColor.HighValue.Value = 80;

                twoColor.LowValue.Color = Color.Gold;
                twoColor.HighValue.Color = Color.Silver;

                var threeColor = wks.ConditionalFormatting.AddThreeColorScale(new ExcelAddress(1, 41, 10, 41));

                var threeIcons = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 43, 10, 43), eExcelconditionalFormatting3IconsSetType.Symbols2);

                var fourIcons = wks.ConditionalFormatting.AddFourIconSet(new ExcelAddress(1, 44, 10, 44), eExcelconditionalFormatting4IconsSetType.RedToBlack);

                var fiveIcons = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress(1, 45, 10, 45), eExcelconditionalFormatting5IconsSetType.Rating);

                var threeGreatherThan = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 48, 10, 48), eExcelconditionalFormatting3IconsSetType.TrafficLights2);

                threeGreatherThan.Icon2.GreaterThanOrEqualTo = false;
                threeGreatherThan.Icon3.GreaterThanOrEqualTo = false;

                wks.Calculate();

                ////ExtLst iconsets are best written last as they will then be read in the correct order
                //var five2 = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress(1, 47, 10, 47), eExcelconditionalFormatting5IconsSetType.Boxes);
                //var threeIcons2 = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 46, 10, 46), eExcelconditionalFormatting3IconsSetType.Triangles);

                pck.SaveAs("C:/epplusTest/Workbooks/conditionalTestEppCopy.xlsx");

                var newPck = new ExcelPackage("C:/epplusTest/Workbooks/conditionalTestEppCopy.xlsx");

                var formattings = newPck.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(formattings[0].Formula, "3");
                Assert.AreEqual(formattings[0].Formula2, "8");
                Assert.AreEqual(formattings[1].Formula, "7");
                Assert.AreEqual(formattings[2].Formula, "1");
                Assert.AreEqual(((IExcelConditionalFormattingContainsText)formattings[3]).ContainText, "o");

                Assert.AreEqual(formattings[4].TimePeriod, eExcelConditionalFormattingTimePeriodType.Last7Days);
                Assert.AreEqual(formattings[5].TimePeriod, eExcelConditionalFormattingTimePeriodType.Yesterday);
                Assert.AreEqual(formattings[6].TimePeriod, eExcelConditionalFormattingTimePeriodType.Today);
                Assert.AreEqual(formattings[7].TimePeriod, eExcelConditionalFormattingTimePeriodType.Tomorrow);
                Assert.AreEqual(formattings[8].TimePeriod, eExcelConditionalFormattingTimePeriodType.LastWeek);
                Assert.AreEqual(formattings[9].TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisWeek);
                Assert.AreEqual(formattings[10].TimePeriod, eExcelConditionalFormattingTimePeriodType.NextWeek);
                Assert.AreEqual(formattings[11].TimePeriod, eExcelConditionalFormattingTimePeriodType.LastMonth);
                Assert.AreEqual(formattings[12].TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisMonth);
                Assert.AreEqual(formattings[13].TimePeriod, eExcelConditionalFormattingTimePeriodType.NextMonth);

                Assert.AreEqual(formattings[14].Type, eExcelConditionalFormattingRuleType.DuplicateValues);

                Assert.AreEqual(formattings[15].Rank, 11);
                Assert.AreEqual(formattings[15].Bottom, false);
                Assert.AreEqual(formattings[15].Percent, false);

                Assert.AreEqual(formattings[16].Rank, 12);
                Assert.AreEqual(formattings[16].Bottom, true);
                Assert.AreEqual(formattings[16].Percent, false);

                Assert.AreEqual(formattings[17].Bottom, false);
                Assert.AreEqual(formattings[17].Percent, true);
                Assert.AreEqual(formattings[17].Rank, 13);

                Assert.AreEqual(formattings[18].Bottom, true);
                Assert.AreEqual(formattings[18].Percent, true);
                Assert.AreEqual(formattings[18].Rank, 14);

                Assert.AreEqual(formattings[19].AboveAverage, true);
                Assert.AreEqual(formattings[19].EqualAverage, false);

                Assert.AreEqual(formattings[20].AboveAverage, true);
                Assert.AreEqual(formattings[20].EqualAverage, true);

                Assert.AreEqual(formattings[21].AboveAverage, false);
                Assert.AreEqual(formattings[21].EqualAverage, false);

                Assert.AreEqual(formattings[22].AboveAverage, false);
                Assert.AreEqual(formattings[22].EqualAverage, true);

                Assert.AreEqual(formattings[23].Type, eExcelConditionalFormattingRuleType.AboveStdDev);
                Assert.AreEqual(formattings[23].StdDev, 1);

                Assert.AreEqual(formattings[24].Type, eExcelConditionalFormattingRuleType.BelowStdDev);
                Assert.AreEqual(formattings[24].StdDev, 2);

                Assert.AreEqual(formattings[25].Type, eExcelConditionalFormattingRuleType.DataBar);
                Assert.AreEqual(formattings[25].As.DataBar.LowValue.Value, 0);
                Assert.AreEqual(formattings[25].As.DataBar.HighValue.Value, 50);

                Assert.AreEqual(formattings[26].Type, eExcelConditionalFormattingRuleType.TwoColorScale);
                Assert.AreEqual(formattings[26].As.TwoColorScale.LowValue.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[26].As.TwoColorScale.HighValue.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[26].As.TwoColorScale.LowValue.Value, 5);
                Assert.AreEqual(formattings[26].As.TwoColorScale.HighValue.Value, 80);
                Assert.AreEqual(formattings[26].As.TwoColorScale.LowValue.Color.ToColorString(), Color.Gold.ToColorString());
                Assert.AreEqual(formattings[26].As.TwoColorScale.HighValue.Color.ToColorString(), Color.Silver.ToColorString());

                Assert.AreEqual(formattings[27].Type, eExcelConditionalFormattingRuleType.ThreeColorScale);
                Assert.AreEqual(formattings[27].As.ThreeColorScale.LowValue.Type, eExcelConditionalFormattingValueObjectType.Min);
                Assert.AreEqual(formattings[27].As.ThreeColorScale.MiddleValue.Type, eExcelConditionalFormattingValueObjectType.Percentile);
                Assert.AreEqual(formattings[27].As.ThreeColorScale.HighValue.Type, eExcelConditionalFormattingValueObjectType.Max);
                Assert.AreEqual(formattings[27].As.ThreeColorScale.MiddleValue.Value, 50);

                Assert.AreEqual(formattings[28].Type, eExcelConditionalFormattingRuleType.ThreeIconSet);
                Assert.AreEqual(formattings[28].As.ThreeIconSet.IconSet, eExcelconditionalFormatting3IconsSetType.Symbols2);
                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);

                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon1.Value, 0);
                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon2.Value, Math.Round(100D / 3, 0));
                Assert.AreEqual(formattings[28].As.ThreeIconSet.Icon3.Value, Math.Round(100D * (2D / 3), 0));

                Assert.AreEqual(formattings[29].Type, eExcelConditionalFormattingRuleType.FourIconSet);
                Assert.AreEqual(formattings[29].As.FourIconSet.IconSet, eExcelconditionalFormatting4IconsSetType.RedToBlack);
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon4.Type, eExcelConditionalFormattingValueObjectType.Percent);

                Assert.AreEqual(formattings[29].As.FourIconSet.Icon1.Value, 0);
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon2.Value, Math.Round(100D / 4, 0));
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon3.Value, Math.Round(100D * (2D / 4), 0));
                Assert.AreEqual(formattings[29].As.FourIconSet.Icon4.Value, 75);

                Assert.AreEqual(formattings[30].Type, eExcelConditionalFormattingRuleType.FiveIconSet);
                Assert.AreEqual(formattings[30].As.FiveIconSet.IconSet, eExcelconditionalFormatting5IconsSetType.Rating);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon4.Type, eExcelConditionalFormattingValueObjectType.Percent);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon5.Type, eExcelConditionalFormattingValueObjectType.Percent);

                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon1.Value, 0);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon2.Value, 20);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon3.Value, 40);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon4.Value, 60);
                Assert.AreEqual(formattings[30].As.FiveIconSet.Icon5.Value, 80);

                Assert.AreEqual(formattings[31].Type, eExcelConditionalFormattingRuleType.ThreeIconSet);
                Assert.AreEqual(formattings[31].As.ThreeIconSet.IconSet, eExcelconditionalFormatting3IconsSetType.TrafficLights2);
                Assert.AreEqual(formattings[31].As.ThreeIconSet.Icon1.GreaterThanOrEqualTo, true);
                Assert.AreEqual(formattings[31].As.ThreeIconSet.Icon2.GreaterThanOrEqualTo, false);
                Assert.AreEqual(formattings[31].As.ThreeIconSet.Icon3.GreaterThanOrEqualTo, false);
            }
        }

        //TODO: We should most likely throw a clearer exception.
        [TestMethod]
        public void CFThrowsIfDatabarValueNotSetOnSave()
        {
            ExcelPackage pck = new ExcelPackage(new MemoryStream());

            var sheet = pck.Workbook.Worksheets.Add("DatabarValueTest");

            var databar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Green);

            databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;

            var stream = new MemoryStream();
            pck.SaveAs(stream);
        }

        [TestMethod]
        public void CFShouldNotThrowIfStyleNotSet()
        {
            //Currently throws bc dxfID. Either give default style or make a better throw.
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var dateFormatting = wks.ConditionalFormatting.AddLast7Days(new ExcelAddress(1, 12, 10, 12));

                MemoryStream stream = new MemoryStream();
                pck.SaveAs(stream);
                var newPck = new ExcelPackage(stream);
                var formattings = newPck.Workbook.Worksheets[0].ConditionalFormatting;
            }
        }

        [TestMethod]
        public void GreaterThanCanReadWrite()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("GreaterThan");

                for (int i = 1; i < 11; i++)
                {
                    ws.Cells[i, 1].Value = i;
                }

                var greaterThanFormatting = ws.ConditionalFormatting.AddGreaterThan(new ExcelAddress(1, 1, 10, 1));
                greaterThanFormatting.Formula = "3";

                greaterThanFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                greaterThanFormatting.Style.Font.Color.Color = Color.Violet;

                pck.Save();

                var readPck = new ExcelPackage(pck.Stream);

                for (int i = 0; i < readPck.Workbook.Worksheets[0].ConditionalFormatting.Count; i++)
                {
                    var format = readPck.Workbook.Worksheets[0].ConditionalFormatting[i];

                    Assert.AreEqual(format.Formula, "3");
                    Assert.AreEqual(Color.Black.ToArgb(), format.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Violet.ToArgb(), format.Style.Font.Color.Color.Value.ToArgb());
                }
            }
        }
        [TestMethod]
        public void CFWholeSheetRangeDeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("$1:$1048576"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }


        [TestMethod]
        public void CFColumnsRangeDeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("$A:$P"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }
        [TestMethod]
        public void CFWholeSheetRange2DeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("A1:XFD1048576"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }

    }
}