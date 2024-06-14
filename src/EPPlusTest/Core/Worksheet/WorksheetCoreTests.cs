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
using System.Drawing;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class WorksheetCoreTests : TestBase
    {
        [TestMethod]
        public void SaveCharToCellShouldBeWrittenAsString()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("CharTest");
                ws.Cells["A1"].Value = 'A';
                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    Assert.AreEqual("A", ws.Cells["A1"].Value);
                }
            }
        }
        [TestMethod]
        public void ValidateAutoFitDontShowHiddenColumns()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AutoFitHidden");
                LoadTestdata(ws);

                ws.Column(2).Hidden = true;

                ws.Cells.AutoFitColumns(); 
                Assert.AreEqual(true, ws.Column(2).Hidden);
                p.Save();
            }
        }

        [TestMethod]
        public void ValidateAutoFitMinWidthRange()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("AutoFitMinWidth");
                LoadTestdata(ws);

                ws.Cells["A:B"].AutoFitColumns(500);
                Assert.AreEqual(500, ws.Column(1).Width);
                Assert.AreEqual(500, ws.Column(2).Width);
                p.Save();
            }
        }

        [TestMethod]
        public void RichTextFlagShouldBeCleanedWhenOverwritingValue()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwriteValue");

                ws.Cells["A1:B2"].RichText.Add("RichText");
                Assert.IsTrue(ws.Cells["A1"].IsRichText);
                ws.Cells["A1"].Value = "Text";
                Assert.IsFalse(ws.Cells["A1"].IsRichText);
                Assert.IsFalse(ws.Cells["A2"].IsRichText);
                Assert.IsFalse(ws.Cells["B1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);
            }
        }

        [TestMethod]
        public void RichTextFlagShouldBeCleanedWhenOverwritingWithArray()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwrite");

                ws.Cells["A1:C3"].RichText.Add("RichText");
                Assert.IsTrue(ws.Cells["A1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);

                ws.Cells["A1:B2"].Value = new string[,] { { "Text", "Text" }, { "Text", "Text" } };
                Assert.IsFalse(ws.Cells["A1"].IsRichText);
                Assert.IsFalse(ws.Cells["B2"].IsRichText);
            }
        }

        [TestMethod]
        public void RichTextColorSettingsShouldUseDifferentInstances()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var rt = ws.Cells["A1"].RichText;
                var rt1 = rt.Add("No1");
                var rt2 = rt.Add("No2");
                rt2.Color = Color.Red;
                Assert.AreEqual(rt1.Color, Color.Empty);
                Assert.AreEqual(rt2.Color, Color.Red);
            }
        }

        [TestMethod]
        public void FormulaShouldBeCleanedWhenOverwritingWithArray()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichTextOverwrite");

                ws.Cells["A1:C3"].FormulaR1C1 = "RC";
                Assert.IsFalse(ws.Cells["A1"].Formula==null);

                ws.Cells["A1:B2"].Value = new string[,] { { "Text", "Text" }, { "Text", "Text" } };
                Assert.IsTrue(ws.Cells["A1"].Formula == "");
            }
        }
        [TestMethod]
        public void AddAutofilterForMergedCells()
        {
            using (var p = OpenPackage("AutofilterMerge.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("AutoFilter");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";
                ws.Cells["C1"].Value = "Col3";
                ws.Cells["A2"].Value = 1;
                ws.Cells["B2"].Value = 10;
                ws.Cells["C2"].Value = 100;
                ws.Cells["A3"].Value = 2;
                ws.Cells["B3"].Value = 20;
                ws.Cells["C3"].Value = 200;
                ws.Cells["A1:B1"].Merge = true;
                ws.Cells["A1:C3"].AutoFilter = true;
                ws.AutoFilter.Columns.AddValueFilterColumn(0);
                ws.AutoFilter.Columns[0].ShowButton = false;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateFirstLastCellTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B4:H10"].Style.Numberformat.Format = "0";
                Assert.IsNull(ws.FirstValueCell);
                Assert.IsNull(ws.LastValueCell);

                ws.Cells["B6:C7"].Value = 1;

                Assert.AreEqual("B6",ws.FirstValueCell.Address);
                Assert.AreEqual("C7", ws.LastValueCell.Address);

                Assert.AreEqual("B6:C7", ws.DimensionByValue.Address);
            }
        }
        [TestMethod]
        public void ValidateDimensionValueLargerRowTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B4:H10"].Style.Numberformat.Format = "0";
                Assert.IsNull(ws.FirstValueCell);
                Assert.IsNull(ws.LastValueCell);

                ws.Cells["D7"].Value = 1;
                ws.Cells["C6"].Value = 1;
                ws.Cells["G6"].Value = 1;
                ws.Cells["D5"].Value = 1;

                Assert.AreEqual("D5", ws.FirstValueCell.Address);
                Assert.AreEqual("D7", ws.LastValueCell.Address);

                Assert.AreEqual("C5:G7", ws.DimensionByValue.Address);
            }
        }
        [TestMethod]
        public void ValidateDimensionValueTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A4:H10"].Style.Numberformat.Format = "0";
                ws.Cells["B6:C7"].Value = 1;


                Assert.AreEqual("B6:C7", ws.DimensionByValue.Address);
            }
        }

    }
}
