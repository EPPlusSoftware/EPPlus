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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
namespace OfficeOpenXml.Core.Range
{
    [TestClass]
    public class RangeAddressTests
    {
        [TestMethod]
        public void MultipleAddressWithWorkbook()
        {
            var split = ExcelAddressBase.SplitFullAddress("'Sheet2'!A:A,A1,[c:\\workbook.xlsx]'Sheet1'!A1");

            Assert.AreEqual(3, split.Count);

            Assert.IsNull(split[0][0]);
            Assert.AreEqual("Sheet2", split[0][1]);
            Assert.AreEqual("A:A", split[0][2]);

            Assert.IsNull(split[1][0]);
            Assert.IsNull(split[1][1]);
            Assert.AreEqual("A1", split[1][2]);

            Assert.AreEqual("c:\\workbook.xlsx", split[2][0]);
            Assert.AreEqual("Sheet1", split[2][1]);
            Assert.AreEqual("A1", split[2][2]);
        }

        [TestMethod]
        public void AddressWithWorkbook()
        {
            var split = ExcelAddressBase.SplitFullAddress("[c:\\workbook.xlsx]'Sheet1'!A1");

            Assert.AreEqual("c:\\workbook.xlsx", split[0][0]);
            Assert.AreEqual("Sheet1", split[0][1]);
            Assert.AreEqual("A1", split[0][2]);
        }
        [TestMethod]
        public void AddressWithWorksheetWithApostrophe()
        {
            var split = ExcelAddressBase.SplitFullAddress("'sheet ''''1'!A1");

            Assert.AreEqual("sheet ''1", split[0][1]);
            Assert.AreEqual("A1", split[0][2]);
        }
        [TestMethod]
        public void AddressWithWorksheetWithoutApostrophe()
        {
            var split = ExcelAddressBase.SplitFullAddress("sheet!A1");

            Assert.AreEqual("sheet", split[0][1]);
            Assert.AreEqual("A1", split[0][2]);
        }

        [TestMethod]
        public void InsertDeleteTest()
        {
            var addr = new ExcelAddressBase("A1:B3");

            Assert.AreEqual(addr.AddRow(2, 4).Address, "A1:B7");
            Assert.AreEqual(addr.AddColumn(2, 4).Address, "A1:F3");
            Assert.AreEqual(addr.DeleteColumn(2, 1).Address, "A1:A3");
            Assert.AreEqual(addr.DeleteRow(2, 2).Address, "A1:B1");

            Assert.AreEqual(addr.DeleteRow(1, 3), null);
            Assert.AreEqual(addr.DeleteColumn(1, 2), null);
        }
        [TestMethod]
        public void SplitAddress()
        {
            var addr = new ExcelAddressBase("C3:F8");

            addr.Insert(new ExcelAddressBase("G9"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("G3"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("C9"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("B2"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("B3"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("D:D"), eShiftTypeInsert.Right);
            addr.Insert(new ExcelAddressBase("5:5"), eShiftTypeInsert.Down);
        }
        [TestMethod]
        public void Addresses()
        {
            var a1 = new ExcelAddress("SalesData!$K$445");
            var a2 = new ExcelAddress("SalesData!$K$445:$M$449,SalesData!$N$448:$Q$454,SalesData!$L$458:$O$464");
            var a3 = new ExcelAddress("SalesData!$K$445:$L$448");
            //var a4 = new ExcelAddress("'[1]Risk]TatTWRForm_TWRWEEKLY20130926090'!$N$527");
            var a5 = new ExcelAddress("Table1[[#All],[Title]]");
            var a6 = new ExcelAddress("Table1[#All]");
            var a7 = new ExcelAddress("Table1[[#Headers],[FirstName]:[LastName]]");
            var a8 = new ExcelAddress("Table1[#Headers]");
            var a9 = new ExcelAddress("Table2[[#All],[SubTotal]]");
            var a10 = new ExcelAddress("Table2[#All]");
            var a11 = new ExcelAddress("Table1[[#All],[Freight]]");
            var a12 = new ExcelAddress("[1]!Table1[[LastName]:[Name]]");
            var a13 = new ExcelAddress("Table1[[#All],[Freight]]");
            var a14 = new ExcelAddress("SalesData!$N$5+'test''1'!$J$33");
        }

        [TestMethod]
        public void IsValidCellAdress()
        {
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("A1"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("A1048576"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("XFD1"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("XFD1048576"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!A1"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!A1048576"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!XFD1"));
            Assert.IsTrue(ExcelCellBase.IsValidCellAddress("Table1!XFD1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1:A1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1:XFD1"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("A1048576:XFD1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("XFD1:XFD1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1:A1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1:XFD1"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!A1048576:XFD1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidCellAddress("Table1!XFD1:XFD1048576"));
        }
        [TestMethod]
        public void IsValidName()
        {
            Assert.IsFalse(ExcelAddressUtil.IsValidName("123sa"));  //invalid start char 
            Assert.IsFalse(ExcelAddressUtil.IsValidName("*d"));     //invalid start char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("\t"));     //invalid start char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("\\t"));    //Backslash at least three chars
            Assert.IsFalse(ExcelAddressUtil.IsValidName("A+1"));   //invalid char
            Assert.IsFalse(ExcelAddressUtil.IsValidName("A%we"));   //Address invalid
            Assert.IsFalse(ExcelAddressUtil.IsValidName("BB73"));   //Address invalid
            Assert.IsTrue(ExcelAddressUtil.IsValidName("BBBB75"));  //Valid
            Assert.IsTrue(ExcelAddressUtil.IsValidName("BB1500005")); //Valid
        }
        [TestMethod]
        public void NamedRangeMovesDownIfRowInsertedAbove()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(1, 1);

                Assert.AreEqual("'NEW'!A3:C4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeDoesNotChangeIfRowInsertedBelow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(4, 1);

                Assert.AreEqual("A2:C3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeExpandsDownIfRowInsertedWithin()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertRow(3, 1);

                Assert.AreEqual("'NEW'!A2:C4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeMovesRightIfColInsertedBefore()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(1, 1);

                Assert.AreEqual("'NEW'!C2:E3", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeUnchangedIfColInsertedAfter()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(5, 1);

                Assert.AreEqual("B2:D3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeExpandsToRightIfColInsertedWithin()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NEW"); 
                var range = sheet.Cells[2, 2, 3, 4];
                var namedRange = sheet.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(5, 1);

                Assert.AreEqual("B2:D3", namedRange.Address);
            }
        }

        [TestMethod]
        public void NamedRangeWithWorkbookScopeIsMovedDownIfRowInsertedAbove()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 1, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet.InsertRow(1, 1);

                Assert.AreEqual("'NEW'!A3:C4", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeWithWorkbookScopeIsMovedRightIfColInsertedBefore()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet = package.Workbook.Worksheets.Add("NEW");
                var range = sheet.Cells[2, 2, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet.InsertColumn(1, 1);

                Assert.AreEqual("'NEW'!C2:D3", namedRange.FullAddress);
            }
        }

        [TestMethod]
        public void NamedRangeIsUnchangedForOutOfScopeSheet()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var sheet2 = package.Workbook.Worksheets.Add("NEW2");
                var range = sheet2.Cells[2, 2, 3, 3];
                var namedRange = workbook.Names.Add("NewNamedRange", range);

                sheet1.InsertColumn(1, 1);

                Assert.AreEqual("B2:C3", namedRange.Address);
            }
        }


        [TestMethod]
        public void ShouldHandleWorksheetSpec()
        {
            var address = "Sheet1!A1:Sheet1!A2";
            var excelAddress = new ExcelAddress(address);
            Assert.AreEqual("Sheet1", excelAddress.WorkSheet);
            Assert.AreEqual(1, excelAddress._fromRow);
            Assert.AreEqual(2, excelAddress._toRow);
        }
        [TestMethod]
        public void IsValidAddress()
        {
            Assert.IsFalse(ExcelCellBase.IsValidAddress("$A12:XY1:3"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("A1$2:XY$13"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("A12$:X$Y$13"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("A12:X$Y$13"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("$A$12:$XY$13,$A12:XY1:3"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("$A$12:"));

            Assert.IsTrue(ExcelCellBase.IsValidAddress("$XFD$1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("$XFE$1048576"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("$XFD$1048577"));

            Assert.IsTrue(ExcelCellBase.IsValidAddress("A12"));
            Assert.IsTrue(ExcelCellBase.IsValidAddress("A$12"));
            Assert.IsTrue(ExcelCellBase.IsValidAddress("$A$12"));
            Assert.IsTrue(ExcelCellBase.IsValidAddress("$A$12:$XY$13"));
            Assert.IsTrue(ExcelCellBase.IsValidAddress("$A$12:$XY$13,$A12:XY$14"));

            Assert.IsFalse(ExcelCellBase.IsValidAddress("$A$12:$XY$13,$A12:XY$14$"));
        }
        [TestMethod]
        public void ClearShouldNotClearSurroundingCells()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Clear");
                ws.Cells[2, 2].Value = "B2";
                ws.Cells[2, 3].Value = "C2";
                ws.Cells[2, 4].Value = "D2";
                ws.Cells[2, 3].Clear();

                Assert.IsNotNull(ws.Cells[2, 2].Value);
                Assert.AreEqual("B2", ws.Cells[2, 2].Value);
                Assert.IsNull(ws.Cells[2, 3].Value);
                Assert.AreEqual("D2", ws.Cells[2, 4].Value);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("CopyRowWise");
                ws.Cells["A1:C3"].Value = 1;
                ws.Cells["D3"].Formula="A1";
                ws.Cells["E3"].Formula = "B2";
                ws.Cells["F3"].Formula = "C3";
                ws.Cells["G3"].Formula = "A$1";
                ws.Cells["H3"].Formula = "B$2";
                ws.Cells["J3"].Formula = "C$3";

                //Validate that formulas are copied correctly row-wise
                ws.Cells["D3"].Copy(ws.Cells["D2"]);
                Assert.AreEqual("#REF!", ws.Cells["D2"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E2"]);
                Assert.AreEqual("B1", ws.Cells["E2"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F2"]);
                Assert.AreEqual("C2", ws.Cells["F2"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G2"]);
                Assert.AreEqual("A$1", ws.Cells["G2"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H2"]);
                Assert.AreEqual("B$2", ws.Cells["H2"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J2"]);
                Assert.AreEqual("C$3", ws.Cells["J2"].Formula);

                ws.Cells["D3"].Copy(ws.Cells["D1"]);
                Assert.AreEqual("#REF!", ws.Cells["D1"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E1"]);
                Assert.AreEqual("#REF!", ws.Cells["E1"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F1"]);
                Assert.AreEqual("C1", ws.Cells["F1"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G1"]);
                Assert.AreEqual("A$1", ws.Cells["G1"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H1"]);
                Assert.AreEqual("B$2", ws.Cells["H1"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J1"]);
                Assert.AreEqual("C$3", ws.Cells["J1"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasColumn()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("CopyColumnWise");
                ws.Cells["A1:C3"].Value = 1;
                ws.Cells["C4"].Formula = "A1";
                ws.Cells["C5"].Formula = "B2";
                ws.Cells["C6"].Formula = "C3";
                ws.Cells["C7"].Formula = "$A1";
                ws.Cells["C8"].Formula = "$B2";
                ws.Cells["C9"].Formula = "$C3";

                //Validate that formulas are copied correctly column-wise
                ws.Cells["C4"].Copy(ws.Cells["B4"]);
                Assert.AreEqual("#REF!", ws.Cells["B4"].Formula);
                ws.Cells["C5"].Copy(ws.Cells["B5"]);
                Assert.AreEqual("A2", ws.Cells["B5"].Formula);
                ws.Cells["C6"].Copy(ws.Cells["B6"]);
                Assert.AreEqual("B3", ws.Cells["B6"].Formula);
                ws.Cells["C7"].Copy(ws.Cells["B7"]);
                Assert.AreEqual("$A1", ws.Cells["B7"].Formula);
                ws.Cells["C8"].Copy(ws.Cells["B8"]);
                Assert.AreEqual("$B2", ws.Cells["B8"].Formula);
                ws.Cells["C9"].Copy(ws.Cells["B9"]);
                Assert.AreEqual("$C3", ws.Cells["B9"].Formula);

                ws.Cells["C4"].Copy(ws.Cells["A4"]);
                Assert.AreEqual("#REF!", ws.Cells["A4"].Formula);
                ws.Cells["C5"].Copy(ws.Cells["A5"]);
                Assert.AreEqual("#REF!", ws.Cells["A5"].Formula);
                ws.Cells["C6"].Copy(ws.Cells["A6"]);
                Assert.AreEqual("A3", ws.Cells["A6"].Formula);
                ws.Cells["C7"].Copy(ws.Cells["A7"]);
                Assert.AreEqual("$A1", ws.Cells["A7"].Formula);
                ws.Cells["C8"].Copy(ws.Cells["A8"]);
                Assert.AreEqual("$B2", ws.Cells["A8"].Formula);
                ws.Cells["C9"].Copy(ws.Cells["A9"]);
                Assert.AreEqual("$C3", ws.Cells["A9"].Formula);
            }
        }
    }
}
