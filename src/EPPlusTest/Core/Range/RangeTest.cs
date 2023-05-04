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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeTest : TestBase
    {
        [TestMethod]
        public void ArrayToCellString()
        {
            var ms = new MemoryStream();
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = new[] { "string1", "string2", "string3" };
                p.Save();
            }
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets["Sheet1"];
                Assert.AreEqual("string1", sheet.Cells[1, 1].Value);
            }
        }

        [TestMethod]
        public void ArrayToCellNull()
        {
            var ms = new MemoryStream();
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = new[] { null, "string2", "string3" };
                p.Save();
            }
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets["Sheet1"];
                Assert.AreEqual(string.Empty, sheet.Cells[1, 1].Value);
            }
        }
        [TestMethod]
        public void ArrayToCellInt()
        {
            var ms = new MemoryStream();
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = new object[] { 1, "string2", "string3" };
                p.Save();
            }
            using (var p = new ExcelPackage(ms))
            {
                var sheet = p.Workbook.Worksheets["Sheet1"];
                Assert.AreEqual(1D, sheet.Cells[1, 1].Value);
            }
        }
        [TestMethod]
        public void ClearRangeWithCommaseparatedAddress()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1:B2, C3:D4"].Value = 5;
                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["Sheet1"];
                    ws.Cells["A1:B2, C3:D4"].Clear();
                    Assert.IsNull(ws.Dimension);
                    p2.Save();
                }

            }
        }
        [TestMethod]
        public void MergeCellsShouldBeSaved()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("MergedCells");

                var r = ws.Cells[1, 1, 1, 5];
                r.Merge = true;
                Assert.AreEqual(1, ws.MergedCells.Count);
                r.Value = "Header";

                Assert.AreEqual(1, ws.MergedCells.Count);
                Assert.AreEqual("A1:E1", ws.MergedCells[0]);
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p.Workbook.Worksheets[0];
                    Assert.AreEqual(1, ws.MergedCells.Count);
                    Assert.AreEqual("A1:E1", ws.MergedCells[0]);
                }
            }
        }
        [TestMethod]
        public void LoadFromCollectionObjectDynamic()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("LoadFromCollection");

                var range = ws.Cells["A1"].LoadFromCollection(new List<object>() { 1, "s", null });
                Assert.AreEqual("A1:A3", range.Address);
                Assert.AreEqual("A1:A3", range.Address);

                range = ws.Cells["B1"].LoadFromCollection(new List<dynamic>() { 1, "s", null });
                Assert.AreEqual("B1:B3", range.Address);

                range = ws.Cells["C1"].LoadFromCollection(new List<dynamic>() { new TestDTO { Name = "Test" } });
                Assert.AreEqual("C1", range.Address);
            }
        }
        [TestMethod]
        public void EncodingCharInFormulaAndValue()
        {
            var textA1 = "\"Hello\vA1\" & \"!\t\nNewLine\"";
            var textB1 = "\"Hello\vB1\" & \"!\t\nNewLine\"";
            using (var p=OpenPackage("EncodeFormula.xlsx",true))
            {
                var ws = p.Workbook.Worksheets.Add("Encoding");
                ws.SetFormula(1, 1, textA1);
                ws.Cells[1, 2].Formula = textB1;
                ws.Calculate();

                Assert.AreEqual(textA1, ws.Cells["A1"].Formula);
                Assert.AreEqual(textB1, ws.GetFormula(1, 2));

                p.Save();
                using(var p2=new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets["Encoding"];
                    Assert.AreEqual(textA1, ws.Cells["A1"].Formula);
                    Assert.AreEqual(textB1, ws.GetFormula(1, 2));
                }
            }
        }
        [TestMethod]
        public void ValidateMergedCell()
        {
            using (var p = OpenPackage("MergeCellsDeleteInsert.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Cells["B2:D2"].Merge = true;
                ws.Cells["B13:D13"].Merge = true;
                ws.Cells["B41:D41"].Merge = true;
                ws.Cells["B42:D42"].Merge = true;
                ws.Cells["B43:D43"].Merge = true;
                ws.Cells["B52:D52, B42:D42, B42:D42"].Merge = true;
                ws.Cells["B52:D52"].Merge = true;
                ws.Cells["B79:D79"].Merge = true;
                ws.Cells["B79:D79"].Merge = true;
                ws.Cells["B102:D102"].Merge = true;
                ws.Cells["B132:D132"].Merge = true;
                ws.Cells["A42:E43"].Delete(eShiftTypeDelete.Up);
                ws.Cells["A42:E44"].Insert(eShiftTypeInsert.Down);
                ws.Cells["A42:E42"].Delete(eShiftTypeDelete.Up);

                foreach (var addr in ws.MergedCells)
                {
                    if (!string.IsNullOrEmpty(addr))
                    {
                        var a = new ExcelAddressBase(addr);
                        for (int r = a._fromRow; r <= a._toRow; r++)
                        {
                            for (int c = a._fromCol; c <= a._toCol; c++)
                            {
                                Assert.IsTrue(ws.Cells[r, c].Merge);
                            }
                        }
                    }
                }

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets["Merge"];
                    ws.Cells["B41:D42"].Merge = true;
                    p2.Save();
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateMergedCellAroundTableShouldNotThrowException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");

                ws.Cells["A1:A2"].Merge = true;
                ws.Cells["A3:E3"].Merge = true;
                ws.Cells["C6:F6"].Merge = true;
                ws.Cells["F3:F5"].Merge = true;
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateMergedCellInsideTableShouldThrowException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");

                ws.Cells["D4:D5"].Merge = true;
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateMergedCellPartlyWithTableShouldThrowException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");

                ws.Cells["D3:D4"].Merge = true;
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ValidateMergedCellEqualTableShouldThrowException()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");

                ws.Cells["D4:E5"].Merge = true;
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ValidateTableAddShouldThrowExceptionMergedCellEqual()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Cells["D4:E5"].Merge = true;
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ValidateTableAddShouldThrowExceptionMergedCellInside()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Cells["D4:D5"].Merge = true;
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ValidateTableAddShouldThrowExceptionMergedCellPartly()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Cells["D3:D4"].Merge = true;
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");
            }
        }
        public void ValidateDeleted()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Merge");
                ws.Cells["D3:D4"].Merge = true;
                ws.Tables.Add(ws.Cells["D4:E5"], "Table1");
            }
        }

    }
}
