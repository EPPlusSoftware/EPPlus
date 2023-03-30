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
using System;

namespace EPPlusTest.Core.Range
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
        public void ShouldHandleWorksheetSpec()
        {
            var address = "Sheet1!A1:Sheet1!A2";
            var excelAddress = new ExcelAddress(address);
            Assert.AreEqual("Sheet1", excelAddress.WorkSheetName);
            Assert.AreEqual(1, excelAddress._fromRow);
            Assert.AreEqual(2, excelAddress._toRow);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AddressWithFullColumnInEndAndCellIsNotValid()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var v = sheet1.Cells["A1:B"]; //Invalid address
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AddressWithFullColumnAtStartAndCellIsNotValid()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var v = sheet1.Cells["A:B1"]; //Invalid address
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AddressWithFullRowInEndAndCellIsNotValid()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var v = sheet1.Cells["A1:2"]; //Invalid address
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AddressWithFullRowAtStartAndCellIsNotValid()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sheet1 = package.Workbook.Worksheets.Add("NEW");
                var v = sheet1.Cells["1:B1"]; //Invalid address
            }
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

            Assert.IsFalse(ExcelCellBase.IsValidAddress("A12:B"));
            Assert.IsFalse(ExcelCellBase.IsValidAddress("A:B12"));
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
        public void VerifyFullAddress()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("AddressVerify");
                Assert.AreEqual("AddressVerify!B6:D8", ws.Cells["B6:D8"].FullAddress);
                Assert.AreEqual("AddressVerify!B6:D8,AddressVerify!B10:D11", ws.Cells["B6:D8,B10:D11"].FullAddress);
            }

        }
        [TestMethod]
        public void VerifyFullWorksheetAddress()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("FullSheet");
                ws1.SetValue("A1", "Col1");
                ws1.SetValue("B1", "Col2");
                ws1.SetValue("A2", 1);
                ws1.SetValue("B2", "Row 1");
                ws1.SetValue("A3", 2);
                ws1.SetValue("B3", "Row 2");
                ws1.SetValue("A4", 3);
                ws1.SetValue("B4", "Row 3");

                var ws2 = pck.Workbook.Worksheets.Add("Formula");
                ws2.SetFormula(1, 1, "VLOOKUP(2,FullSheet,2,FALSE)");
                ws2.SetFormula(2, 1, "VLOOKUP(3,'FullSheet',2,FALSE)");
                ws2.Calculate();
                Assert.AreEqual("Row 2", ws2.Cells["A1"].Value);
                Assert.AreEqual("Row 3", ws2.Cells["A2"].Value);
            }
        }
        [TestMethod]
        public void VerifyFullWorksheetAddressR1C1Start()
        {
            using (var pck = new ExcelPackage())
            {
                var wb = pck.Workbook;

                var ws = wb.Worksheets.Add("RC01");
                var n=wb.Names.Add("Name1", ws.Cells["A1"]);
                Assert.AreEqual("'RC01'!$A$1", n.FullAddressAbsolute);

                ws = wb.Worksheets.Add("CR01");
                n = wb.Names.Add("Name2", ws.Cells["A1"]);
                Assert.AreEqual("CR01!$A$1", n.FullAddressAbsolute);

                ws = wb.Worksheets.Add("C1");
                n = wb.Names.Add("Name3", ws.Cells["A1"]);
                Assert.AreEqual("'C1'!$A$1", n.FullAddressAbsolute);

                ws = wb.Worksheets.Add("C0001");
                n = wb.Names.Add("Name3", ws.Cells["A1"]);
                Assert.AreEqual("'C0001'!$A$1", n.FullAddressAbsolute);

                ws = wb.Worksheets.Add("r9");
                n = wb.Names.Add("Name4", ws.Cells["A1"]);
                Assert.AreEqual("'r9'!$A$1", n.FullAddressAbsolute);

                ws = wb.Worksheets.Add("R009_cc");
                n = wb.Names.Add("Name4", ws.Cells["A1"]);
                Assert.AreEqual("'R009_cc'!$A$1", n.FullAddressAbsolute);
            }
        }
        [TestMethod]
        public void R1C1OffSheetReferenceShouldNotBeASharedFormula()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");
                ExcelWorksheet ws2 = package.Workbook.Worksheets.Add("Sheet2");
                ws2.Cells[1, 3, 4, 3].Value = 2;
                ws2.Cells[2, 3].Value = 1;
                ws2.Cells[4, 3].Value = 4;
                ws.Cells[1, 2].Value = 1;
                ws.Cells[2, 2].Value = 2;
                ws.Cells[3, 2].Value = 3;
                ws.Cells[4, 2].Value = 4;
                ws.Cells[1, 1, 4, 1].FormulaR1C1 = "COUNTIF(Sheet2!R1C3:R4C3, RC2)";
                Assert.AreEqual(0, ws._sharedFormulas.Count);
            }
        }
        [TestMethod]
        public void R1C1ReferenceShouldBeASharedFormula()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 3, 4, 3].Value = 2;
                ws.Cells[2, 3].Value = 1;
                ws.Cells[4, 3].Value = 4;
                ws.Cells[1, 2].Value = 1;
                ws.Cells[2, 2].Value = 2;
                ws.Cells[3, 2].Value = 3;
                ws.Cells[4, 2].Value = 4;
                ws.Cells["A10:A14"].FormulaR1C1 = "COUNTIF(R1C3:R4C3, RC2)";
                Assert.AreEqual(1, ws._sharedFormulas.Count);
            }
        }
        [TestMethod]
        public void ValidateMultiAddress_b1_b5_c4()
        {
            var a1 = new ExcelAddress("b1:b5:c4");

            Assert.AreEqual(1, a1.Start.Row);
            Assert.AreEqual(2, a1.Start.Column);
            Assert.AreEqual(5, a1.End.Row);
            Assert.AreEqual(3, a1.End.Column);
        }
        [TestMethod]
        public void ValidateMultiAddress_b4_b1_c4()
        {
            var a1 = new ExcelAddress("b4:b3:c5");

            Assert.AreEqual(3, a1.Start.Row);
            Assert.AreEqual(2, a1.Start.Column);
            Assert.AreEqual(5, a1.End.Row);
            Assert.AreEqual(3, a1.End.Column);
        }
        [TestMethod]
        public void ValidateMultiAddress_F5_G2_F7_G8()
        {
            var a1 = new ExcelAddress("F5:G2:F7:G8");

            Assert.AreEqual(2, a1.Start.Row);
            Assert.AreEqual(6, a1.Start.Column);
            Assert.AreEqual(8, a1.End.Row);
            Assert.AreEqual(7, a1.End.Column);
        }
        [TestMethod]
        public void ValidateMultiAddress_G8_F7_G2_F5()
        {
            var a1 = new ExcelAddress("G8:F7:G2:F5");

            Assert.AreEqual(2, a1.Start.Row);
            Assert.AreEqual(6, a1.Start.Column);
            Assert.AreEqual(8, a1.End.Row);
            Assert.AreEqual(7, a1.End.Column);
        }
        [TestMethod]
        public void ValidateInsertIntoWorksheetNameWithApostrophe()
        {
            ExcelAddressBase a1 = new ExcelAddressBase("'d''étude'!A1");

            var a2 = a1.AddRow(1, 1).AddColumn(1,1);
            Assert.AreEqual("'d''étude'!B2", a2.Address);
        }
        [TestMethod]
        public void ValidateAddressWithDoubleWorksheetReferences()
        {
            ExcelAddressBase a1 = new ExcelAddressBase("'a'!a1:'a'!A3");

            Assert.AreEqual(1, a1._fromRow);
            Assert.AreEqual(1, a1._fromCol);
            Assert.AreEqual(3, a1._toRow);
            Assert.AreEqual(1, a1._toCol);
            Assert.AreEqual("A1:A3", a1.LocalAddress);
        }
        [TestMethod]
        public void ValidateLocalAddressWhenMultiAddress()
        {
            ExcelAddressBase a1 = new ExcelAddressBase("'a'!a1,'b'!A3");

            Assert.AreEqual("A1,A3", a1.LocalAddress);
        }

        [TestMethod]
        public void ValidateGetWorksheetPart()
        {
            int ix = 0;
            ExcelAddressBase.GetWorksheetPart("A1:A2", "Sheet1", ref ix);
            Assert.AreEqual(0, ix);

            ExcelAddressBase.GetWorksheetPart("sheet1!A1:A2", "Sheet1", ref ix);
            Assert.AreEqual(7, ix);
            ExcelAddressBase.GetWorksheetPart("'sheet 1'!A1:A2", "Sheet1", ref ix);
            Assert.AreEqual(10, ix);
        }
        [TestMethod]
        public void IntersectReversedTopLeftTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("A1:C3"));
            Assert.AreEqual(2, ia1.Addresses.Count);
            Assert.AreEqual("D2:D4", ia1.Addresses[0].Address);
            Assert.AreEqual("B4:C4", ia1.Addresses[1].Address);

        }
        [TestMethod]
        public void IntersectReversedTopRightTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("C1:E3"));
            Assert.AreEqual(2, ia1.Addresses.Count);
            Assert.AreEqual("B2:B4", ia1.Addresses[0].Address);
            Assert.AreEqual("C4:D4", ia1.Addresses[1].Address);

        }

        [TestMethod]
        public void IntersectReversedTopTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("A2:D3"));
            Assert.AreEqual("B4:D4", ia1.Address);
        }
        [TestMethod]
        public void IntersectReversedLeftTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("A2:C4"));
            Assert.AreEqual("D2:D4", ia1.Address);
        }
        [TestMethod]
        public void IntersectReversedOutSideTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("B2:D4"));
            Assert.IsNull(ia1);
            var ia2 = address.IntersectReversed(new ExcelAddressBase("A1:E5"));
            Assert.IsNull(ia2);
        }
        [TestMethod]
        public void IntersectReversedInsideTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("C3"));
            Assert.AreEqual(4, ia1.Addresses.Count);
            Assert.AreEqual("B2:B4", ia1.Addresses[0].Address);
            Assert.AreEqual("C2:D2", ia1.Addresses[1].Address);
            Assert.AreEqual("D3:D4", ia1.Addresses[2].Address);
            Assert.AreEqual("C4", ia1.Addresses[3].Address);
        }

        [TestMethod]
        public void IntersectReversedBottomLeftTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("A4:B5"));
            Assert.AreEqual(2, ia1.Addresses.Count);
            Assert.AreEqual("B2:D3", ia1.Addresses[0].Address);
            Assert.AreEqual("C4:D4", ia1.Addresses[1].Address);

        }
        [TestMethod]
        public void IntersectReversedBottomRightTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("C3:E12"));
            Assert.AreEqual(2, ia1.Addresses.Count);
            Assert.AreEqual("B2:B4", ia1.Addresses[0].Address);
            Assert.AreEqual("C2:D2", ia1.Addresses[1].Address);
        }

        [TestMethod]
        public void IntersectReversedBottomTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("A4:E5"));
            Assert.AreEqual("B2:D3", ia1.Address);
        }
        [TestMethod]
        public void IntersectReversedRightTest()
        {
            var address = new ExcelAddressBase("B2:D4");

            var ia1 = address.IntersectReversed(new ExcelAddressBase("C2:E4"));
            Assert.AreEqual("B2:B4", ia1.Address);
        }

    }
}