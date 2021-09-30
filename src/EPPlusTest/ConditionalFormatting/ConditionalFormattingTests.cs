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
using System.IO;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Drawing;

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
        public static void MyClassCleanup()
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
        public void DatabarChangingAddressAddsConditionalFormatNodeInSchemaOrder()
        {
            var ws = _pck.Workbook.Worksheets.Add("DatabarAddressing");
            // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
            ws.HeaderFooter.AlignWithMargins = true;
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
            Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
            cf.Address = new ExcelAddress("C3");
            Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
            Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
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
                var cf=ws.ConditionalFormatting.AddBetween(ws.Cells["A1:A3"]);
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
                    Assert.AreEqual(Color.Red.ToArgb(),cf.Style.Font.Color.Color.Value.ToArgb());
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
    }
}