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
    public class TableTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Table.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void TableWithSubtotalsParensInColumnName()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSubtotParensColumnName");
            ws.Cells["B2"].Value = "Header 1";
            ws.Cells["C2"].Value = "Header (2)";
            ws.Cells["B3"].Value = 1;
            ws.Cells["B4"].Value = 2;
            ws.Cells["C3"].Value = 3;
            ws.Cells["C4"].Value = 4;
            var table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTableParamHeader");
            table.ShowTotal = true;
            table.ShowHeader = true;
            table.Columns[0].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
            table.Columns[1].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
            ws.Cells["B5"].Calculate();
            Assert.AreEqual(3.0, ws.Cells["B5"].Value);
            ws.Cells["C5"].Calculate();
            Assert.AreEqual(7.0, ws.Cells["C5"].Value);
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestTableNameCanNotStartsWithNumber()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Table");
                var tbl = ws.Tables.Add(ws.Cells["A1"], "5TestTable");
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestTableNameCanNotContainWhiteSpaces()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Table");
                var tbl = ws.Tables.Add(ws.Cells["A1"], "Test Table");
            }
        }

        [TestMethod]
        public void TestTableNameCanStartsWithBackSlash()
        {
            var ws = _pck.Workbook.Worksheets.Add("NameStartWithBackSlash");
            var tbl = ws.Tables.Add(ws.Cells["A1"], "\\TestTable");
        }

        [TestMethod]
        public void TestTableNameCanStartsWithUnderscore()
        {
            var ws = _pck.Workbook.Worksheets.Add("NameStartWithUnderscore");
            var tbl = ws.Tables.Add(ws.Cells["A1"], "_TestTable");
        }
        [TestMethod]
        public void TableTotalsRowFunctionEscapesSpecialCharactersInColumnName()
        {
            var ws = _pck.Workbook.Worksheets.Add("TotalsFormulaTest");
            ws.Cells["A1"].Value = "Col1";
            ws.Cells["B1"].Value = "[#'Col2']";
            var tbl = ws.Tables.Add(ws.Cells["A1:B2"], "Table1");
            tbl.ShowTotal = true;
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            Assert.AreEqual("SUBTOTAL(109,Table1['['#''Col2''']])", ws.Cells["B3"].Formula);
        }
    }
}
