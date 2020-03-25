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
using OfficeOpenXml.Table;
using System;
namespace EPPlusTest.Table
{
    [TestClass]
    public class TableInsertTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("TableInsert.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void TableInsertTop()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("TableInsertTop");
            LoadTestdata(ws, 100);

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertTop");
            ws.Cells["A102"].Value = "Shift Me Down";
            tbl.InsertRow(0);

            Assert.AreEqual("A1:D101", tbl.Address.Address);
            Assert.IsNull(tbl.Range.Offset(1, 0, 1, 1).Value);
            Assert.IsNull(tbl.Range.Offset(1, 1, 1, 1).Value);
            Assert.IsNull(tbl.Range.Offset(1, 2, 1, 1).Value);
            Assert.IsNull(tbl.Range.Offset(1, 3, 1, 1).Value);
            Assert.IsNull(tbl.Range.Offset(1, 4, 1, 1).Value);
            Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
            tbl.InsertRow(0, 3);
            Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
        }
        [TestMethod]
        public void TableInsertBottom()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("TableInsertBottom");
            LoadTestdata(ws, 100);
            ws.Cells["A102"].Value = "Shift Me Down";
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertBottom");
            tbl.AddRow(1);
            Assert.AreEqual("A1:D101", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
            tbl.AddRow(3);
            Assert.AreEqual("A1:D104", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
        }
        [TestMethod]
        public void TableInsertBottomWithTotal()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("TableInsertBottomTotal");
            LoadTestdata(ws, 100);
            ws.Cells["A102"].Value = "Shift Me Down";
            ws.Cells["E5"].Value = "Don't Shift Me";

            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertBottomTotal");
            tbl.ShowTotal = true;
            tbl.Columns[0].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Count;
            tbl.Columns[2].TotalsRowFunction = RowFunctions.Average;
            tbl.Columns[3].TotalsRowFunction = RowFunctions.CountNums;
            tbl.AddRow(1);
            Assert.AreEqual("A1:D102", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
            tbl.AddRow(3);
            Assert.AreEqual("A1:D105", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
            Assert.AreEqual("Don't Shift Me", ws.Cells["E5"].Value);
        }
        [TestMethod]
        public void TableInsertInside()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("TableInsertInside");
            LoadTestdata(ws, 100);
            ws.Cells["A102"].Value = "Shift Me Down";
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableInsertInside");
            tbl.InsertRow(98);
            Assert.AreEqual("A1:D101", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A103"].Value);
            tbl.InsertRow(1, 3);
            Assert.AreEqual("A1:D104", tbl.Address.Address);
            Assert.AreEqual("Shift Me Down", ws.Cells["A106"].Value);
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TablePositionNegative()
        {
            //Setup
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Table1");
                var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "Table1");
                tbl.InsertRow(-1);
            }
        }
        [TestMethod]
        public void TableAddRowToMax()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("TableMaxRow");
            LoadTestdata(ws, 100);
            var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableMaxRow");
            //Act
            tbl.AddRow(ExcelPackage.MaxRows - 100);
            //Assert
            Assert.AreEqual(ExcelPackage.MaxRows, tbl.Address._toRow);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void TableAddRowOverMax()
        {
            using (var p = new ExcelPackage())
            {
                //Setup
                var ws = p.Workbook.Worksheets.Add("TableOverMaxRow");
                LoadTestdata(ws, 100);
                var tbl = ws.Tables.Add(ws.Cells["A1:D100"], "TableOverMaxRow");
                //Act
                tbl.AddRow(ExcelPackage.MaxRows - 99);
                //Assert
                Assert.AreEqual(ExcelPackage.MaxRows, tbl.Address._toRow);
            }
        }

    }
}
