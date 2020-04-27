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
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Drawing;

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
                var ws = pck.Workbook.Worksheets.Add("TableNoWhiteSpace");
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
            var tbl = ws.Tables.Add(ws.Cells["A1:B2"], "TableFormulaTest");
            tbl.ShowTotal = true;
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            Assert.AreEqual("SUBTOTAL(109,TableFormulaTest['['#''Col2''']])", ws.Cells["B3"].Formula);
        }
        [TestMethod]
        public void ValidateEncodingForTableColumnNames()
        {
            var ws = _pck.Workbook.Worksheets.Add("ValidateTblColumnNames");
            ws.Cells["A1"].Value = "Col1>";
            ws.Cells["B1"].Value = "Col1&gt;";
            var tbl = ws.Tables.Add(ws.Cells["A1:C2"], "TableValColNames");
            Assert.AreEqual("Col1>",tbl.Columns[0].Name);
            Assert.AreEqual("Col1&gt;", tbl.Columns[1].Name);
            Assert.AreEqual("Column3", tbl.Columns[2].Name);
        }
        [TestMethod]
        public void TableTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("Table");
            ws.Cells["B1"].Value = 123;
            var tbl = ws.Tables.Add(ws.Cells["B1:P12"], "TestTable");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Custom;

            tbl.ShowFirstColumn = true;
            tbl.ShowTotal = true;
            tbl.ShowHeader = true;
            tbl.ShowLastColumn = true;
            tbl.ShowFilter = false;
            Assert.AreEqual(tbl.ShowFilter, false);
            ws.Cells["K2"].Value = 5;
            ws.Cells["J3"].Value = 4;

            tbl.Columns[8].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
            tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])", tbl.Columns[9].Name);
            tbl.Columns[14].CalculatedColumnFormula = "TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]";
            ws.Cells["B2"].Value = 1;
            ws.Cells["B3"].Value = 2;
            ws.Cells["B4"].Value = 3;
            ws.Cells["B5"].Value = 4;
            ws.Cells["B6"].Value = 5;
            ws.Cells["B7"].Value = 6;
            ws.Cells["B8"].Value = 7;
            ws.Cells["B9"].Value = 8;
            ws.Cells["B10"].Value = 9;
            ws.Cells["B11"].Value = 10;
            ws.Cells["B12"].Value = 11;
            ws.Cells["C7"].Value = "Table test";
            ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

            tbl = ws.Tables.Add(ws.Cells["a12:a13"], "");

            tbl = ws.Tables.Add(ws.Cells["C16:Y35"], "");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;
            tbl.ShowFirstColumn = true;
            tbl.ShowLastColumn = true;
            tbl.ShowColumnStripes = true;
            Assert.AreEqual(tbl.ShowFilter, true);
            tbl.Columns[2].Name = "Test Column Name";

            ws.Cells["G50"].Value = "Timespan";
            ws.Cells["G51"].Value = new DateTime(new TimeSpan(1, 1, 10).Ticks); //new DateTime(1899, 12, 30, 1, 1, 10);
            ws.Cells["G52"].Value = new DateTime(1899, 12, 30, 2, 3, 10);
            ws.Cells["G53"].Value = new DateTime(1899, 12, 30, 3, 4, 10);
            ws.Cells["G54"].Value = new DateTime(1899, 12, 30, 4, 5, 10);

            ws.Cells["G51:G55"].Style.Numberformat.Format = "HH:MM:SS";
            tbl = ws.Tables.Add(ws.Cells["G50:G54"], "");
            tbl.ShowTotal = true;
            tbl.ShowFilter = false;
            tbl.Columns[0].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
        }

        [TestMethod]
        public void TableDeleteTest()
        {
                var wb = _pck.Workbook;
                var sheets = new[]
                {
                    wb.Worksheets.Add("WorkSheet A"),
                    wb.Worksheets.Add("WorkSheet B")
                };
                for (int i = 1; i <= 4; i++)
                {
                    var cell = sheets[0].Cells[1, i];
                    cell.Value = cell.Address + "_";
                    cell = sheets[1].Cells[1, i];
                    cell.Value = cell.Address + "_";
                }

                for (int i = 6; i <= 11; i++)
                {
                    var cell = sheets[0].Cells[3, i];
                    cell.Value = cell.Address + "_";
                    cell = sheets[1].Cells[3, i];
                    cell.Value = cell.Address + "_";
                }
                var tables = new[]
                {
                    sheets[1].Tables.Add(sheets[1].Cells["A1:D73"], "TableDeletea"),
                    sheets[0].Tables.Add(sheets[0].Cells["A1:D73"], "TableDelete2"),
                    sheets[1].Tables.Add(sheets[1].Cells["F3:K10"], "TableDeleteb"),
                    sheets[0].Tables.Add(sheets[0].Cells["F3:K10"], "TableDelete3"),
                };
                Assert.AreEqual(5, wb._nextTableID);
                Assert.AreEqual(1, tables[0].Id);
                Assert.AreEqual(2, tables[1].Id);
                try
                {
                    sheets[0].Tables.Delete("TableDeletea");
                    Assert.Fail("ArgumentException should have been thrown.");
                }
                catch (ArgumentOutOfRangeException) { }
                sheets[1].Tables.Delete("TableDeletea");
                Assert.AreEqual(1, tables[1].Id);
                Assert.AreEqual(2, tables[2].Id);

                try
                {
                    sheets[1].Tables.Delete(4);
                    Assert.Fail("ArgumentException should have been thrown.");
                }
                catch (ArgumentOutOfRangeException) { }
                var range = sheets[0].Cells[sheets[0].Tables[1].Address.Address];
                sheets[0].Tables.Delete(1, true);
                foreach (var cell in range)
                {
                    Assert.IsNull(cell.Value);
                }
        }
    }
}
