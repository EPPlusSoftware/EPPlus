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
using OfficeOpenXml.Drawing;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;

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
            var table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTableParathesesHeader");
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
        public void TableWithSubtotalsBracketInColumnName()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableSubtotBracketColumnName");
            ws.Cells["B2"].Value = "Header 1 & 7";
            ws.Cells["C2"].Value = "Header [test]";
            ws.Cells["B3"].Value = 1;
            ws.Cells["B4"].Value = 2;
            ws.Cells["C3"].Value = 3;
            ws.Cells["C4"].Value = 4;
            var table = ws.Tables.Add(ws.Cells["B2:C4"], "TestTableBracketHeader");
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
            Assert.AreEqual("Col1>", tbl.Columns[0].Name);
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
            ws.Cells["O5"].Value = 11;
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
            using (var p = OpenPackage("TableDeleteTest.xlsx", true))
            {
                var wb = p.Workbook;
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
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void DeleteTablesFromTemplate()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Tablews1");
                ws.Tables.Add(new ExcelAddressBase("A1:C3"), "Table1");
                ws.Tables.Add(new ExcelAddressBase("D1:G7"), "Table2");

                Assert.AreEqual(2, ws.Tables.Count);
                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    Assert.AreEqual(2, ws.Tables.Count);
                    ws.Tables.Delete(0);
                    ws.Tables.Delete("Table2");

                    Assert.AreEqual(0, ws.Tables.Count);
                    p2.Save();
                    using (var p3 = new ExcelPackage(p2.Stream))
                    {
                        Assert.AreEqual(0, p3.Workbook.Worksheets[0].Tables.Count);
                    }
                }
            }
        }
        [TestMethod]
        public void ValidateTableSaveLoad()
        {
            using (var p1 = OpenPackage("table.xlsx", true))
            {
                var sheet = p1.Workbook.Worksheets.Add("Tables");

                // headers
                sheet.Cells["A1"].Value = "Month";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["C1"].Value = "VAT";
                sheet.Cells["D1"].Value = "Total";

                var rnd = new Random();
                for (var row = 2; row < 12; row++)
                {
                    sheet.Cells[row, 1].Value = new DateTimeFormatInfo().GetMonthName(row);
                    sheet.Cells[row, 2].Value = rnd.Next(10000, 100000);
                    sheet.Cells[row, 3].Formula = $"B{row} * 0.25";
                    sheet.Cells[row, 4].Formula = $"B{row} + C{row}";
                }
                sheet.Cells["B2:D13"].Style.Numberformat.Format = "€#,##0.00";

                var range = sheet.Cells["A1:D11"];

                // create the table
                var table = sheet.Tables.Add(range, "myTable");
                // configure the table
                table.ShowHeader = true;
                table.ShowFirstColumn = true;
                table.TableStyle = TableStyles.Dark2;
                // add a totals row under the data
                table.ShowTotal = true;
                table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[3].TotalsRowFunction = RowFunctions.Sum;

                // Calculate all the formulas including the totals row.
                // This will give input to the AutofitColumns call
                range.Calculate();
                range.AutoFitColumns();

                p1.Save();
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    sheet = p2.Workbook.Worksheets["Tables"];
                    // get a table by its name and change properties
                    var myTable = sheet.Tables["myTable"];
                    myTable.TableStyle = TableStyles.Medium8;
                    myTable.ShowFirstColumn = false;
                    myTable.ShowLastColumn = true;
                    Assert.AreEqual(TableStyles.Medium8, myTable.TableStyle);
                    SaveWorkbook("Table2.xlsx", p2);
                    using (var p3 = new ExcelPackage(p2.Stream))
                    {
                        sheet = p3.Workbook.Worksheets["Tables"];
                        // get a table by its name and change properties
                        sheet.Tables.Delete("myTable");

                        SaveWorkbook("Table3.xlsx", p3);
                    }
                }
            }
        }
        [TestMethod]
        public void AddRowShouldAdjustSubtotals()
        {
            using (var package = OpenPackage("TableAdjustSubtotals.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Tables");

                // headers
                sheet.Cells["A1"].Value = "Month";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["C1"].Value = "VAT";
                sheet.Cells["D1"].Value = "Total";

                var rnd = new Random();
                for (var row = 2; row < 12; row++)
                {
                    sheet.Cells[row, 1].Value = new DateTimeFormatInfo().GetMonthName(row);
                    sheet.Cells[row, 2].Value = rnd.Next(10000, 100000);
                    sheet.Cells[row, 3].Formula = $"B{row} * 0.25";
                    sheet.Cells[row, 4].Formula = $"B{row} + C{row}";
                }
                sheet.Cells["B2:D13"].Style.Numberformat.Format = "€#,##0.00";

                var range = sheet.Cells["A1:D11"];

                // create the table
                var table = sheet.Tables.Add(range, "myTable");
                // configure the table
                table.ShowHeader = true;
                table.ShowFirstColumn = true;
                table.ShowFilter = false;
                table.TableStyle = TableStyles.Dark2;
                // add a totals row under the data
                table.ShowTotal = true;
                table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[3].TotalsRowFunction = RowFunctions.Sum;

                // insert rows
                var rowRange = table.AddRow();
                var newRowIx = rowRange.Start.Row;
                sheet.Cells[newRowIx, 1].Value = new DateTimeFormatInfo().GetMonthName(newRowIx);
                sheet.Cells[newRowIx, 2].Value = rnd.Next(10000, 100000);
                sheet.Cells[newRowIx, 3].Formula = $"B{newRowIx} * 0.25";
                sheet.Cells[newRowIx, 4].Formula = $"B{newRowIx} + C{newRowIx}";

                // Calculate all the formulas including the totals row.
                sheet.Calculate();
                sheet.Cells.AutoFitColumns();

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void ValidateCalculatedColumn()
        {
            using (var package = OpenPackage("TableCalculatedColumn.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Tables");

                // headers
                sheet.Cells["C1"].Value = "Month";
                sheet.Cells["D1"].Value = "Sales";
                sheet.Cells["E1"].Value = "VAT";
                sheet.Cells["F1"].Value = "Total";
                sheet.Cells["G1"].Value = "Formula";

                var rnd = new Random();
                for (var row = 2; row < 12; row++)
                {
                    sheet.Cells[row, 3].Value = new DateTimeFormatInfo().GetMonthName(row);
                    sheet.Cells[row, 4].Value = rnd.Next(10000, 100000);
                    sheet.Cells[row, 5].Formula = $"D{row} * 0.25";
                    sheet.Cells[row, 6].Formula = $"D{row} + E{row}";
                }
                sheet.Cells["D2:G13"].Style.Numberformat.Format = "€#,##0.00";

                var range = sheet.Cells["C1:G11"];

                // create the table
                var table = sheet.Tables.Add(range, "myTable");
                // configure the table
                table.ShowHeader = true;
                table.ShowTotal = true;

                var formula = "mytable[[#this row],[Sales]]+mytable[[#this row],[VAT]]";
                table.Columns[4].CalculatedColumnFormula = formula;
                
                //Assert
                Assert.AreEqual(formula, table.Columns[4].CalculatedColumnFormula);
                Assert.AreEqual(formula, sheet.Cells["G2"].Formula);
                Assert.AreEqual(formula, sheet.Cells["G3"].Formula);
                Assert.AreEqual(formula, sheet.Cells["G11"].Formula);

                table.AddRow(3);
                Assert.AreEqual(formula, sheet.Cells["G13"].Formula);


                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void RenameTableWithCalculatedColumnFormulas()
        {
            using (var p = new ExcelPackage())
            {
                // Get the worksheet containing the tables
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");

                // Get the tables and check the calculated column formulas
                var tbl1 = ws1.Tables.Add(ws1.Cells["A1:C2"], "Table1");
                tbl1.Columns[2].CalculatedColumnFormula = "Table1[Column1]+Table1[Column2]";

                var tbl2 = ws1.Tables.Add(ws1.Cells["E1:G2"], "Table2");
                tbl2.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]";

                ws2.SetFormula(1, 1, "Table1[[#This Row],[Column1]]");
                ws2.Cells["B1:B2"].Formula = "Table1[[#This Row],[Column3]]";
                p.Workbook.Names.AddFormula("TableRef", "Table1[[#This Row],[Column1]]");
                Assert.AreEqual("Table1[Column1]+Table1[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("Table1[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]", tbl2.Columns["Column3"].CalculatedColumnFormula);

                // Rename Table1 to Table3 and check the formulas were updated
                tbl1.Name = "NewTableName";
                Assert.AreEqual("NewTableName[Column1]+NewTableName[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("NewTableName[[#This Row],[Column1]]+Table2[[#This Row],[Column2]]", tbl2.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("NewTableName[[#This Row],[Column1]]", p.Workbook.Worksheets[1].Cells["A1"].Formula);
                Assert.AreEqual("NewTableName[[#This Row],[Column3]]", p.Workbook.Worksheets[1].Cells["B2"].Formula);
                Assert.AreEqual("NewTableName[[#This Row],[Column1]]", p.Workbook.Names["TableRef"].Formula);
            }
        }
        [TestMethod]
        public void RenameTableWithCalculatedColumnFormulasSameStartOfTableName()
        {
            using (var p = new ExcelPackage())
            {
                // Create some worksheets
                var ws1 = p.Workbook.Worksheets.Add("Sheet1");
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");

                // Create some tables with calculated column formulas
                var tbl1 = ws1.Tables.Add(ws1.Cells["A1:C2"], "Table1");
                tbl1.Columns[2].CalculatedColumnFormula = "Table1[Column1]+Table1[Column2]";

                var tbl2 = ws1.Tables.Add(ws1.Cells["E1:G2"], "Table12");
                tbl2.Columns[2].CalculatedColumnFormula = "Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]";

                // Create some references outside of the table
                ws2.SetFormula(1, 1, "Table1[[#This Row],[Column1]]");
                ws2.Cells["B1:B2"].Formula = "Table1[[#This Row],[Column3]]";
                p.Workbook.Names.AddFormula("TableRef", "Table1[[#This Row],[Column1]]");
                Assert.AreEqual("Table1[Column1]+Table1[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", tbl2.Columns["Column3"].CalculatedColumnFormula);
                Assert.AreEqual("Table1[Column1]+Table1[Column2]", ws1.Cells["C2"].Formula);
                Assert.AreEqual("Table1[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", ws1.Cells["G2"].Formula);

                // Rename Table1 to Table3 and check the formulas were updated
                tbl1.Name = "Table3";
                Assert.AreEqual("Table3[Column1]+Table3[Column2]", tbl1.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("Table3[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", tbl2.Columns[2].CalculatedColumnFormula);
                Assert.AreEqual("Table3[Column1]+Table3[Column2]", ws1.Cells["C2"].Formula);
                Assert.AreEqual("Table3[[#This Row],[Column1]]+Table12[[#This Row],[Column2]]", ws1.Cells["G2"].Formula);
                Assert.AreEqual("Table3[[#This Row],[Column1]]", p.Workbook.Worksheets[1].Cells["A1"].Formula);
                Assert.AreEqual("Table3[[#This Row],[Column3]]", p.Workbook.Worksheets[1].Cells["B2"].Formula);
                Assert.AreEqual("Table3[[#This Row],[Column1]]", p.Workbook.Names["TableRef"].Formula);
            }
        }
        [TestMethod]
        public void CalculatedColumnFormula_SetToEmptyString()
        {
            using (var pck = new ExcelPackage())
            {
                // Set up a worksheet containing a table
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                wks.Cells["A1"].Value = "Col1";
                wks.Cells["B1"].Value = "Col2";
                wks.Cells["C1"].Value = "Col3";
                wks.Cells["A2"].Value = 1;
                wks.Cells["B2"].Value = 2;
                var table1 = wks.Tables.Add(wks.Cells["A1:C2"], "Table1");
                var formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
                table1.Columns[2].CalculatedColumnFormula = formula;

                // Check the calculated column formula
                Assert.AreEqual(formula, wks.Cells["C2"].Formula);
                Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

                // Remove the calculated column formula from the table
                table1.Columns["Col3"].CalculatedColumnFormula = null;

                // Check the formula has been removed from the table
                Assert.IsTrue(string.IsNullOrEmpty(wks.Cells["C2"].Formula));
                Assert.IsTrue(string.IsNullOrEmpty(table1.Columns["Col3"].CalculatedColumnFormula));

                pck.SaveAs(@"C:\epplusTest\Testoutput\CalculatedColumnFormula_SetToEmptyString.xlsx");

                // NOW OPEN THE FILE IN EXCEL - IS IT CORRUPT?
                Assert.Inconclusive();
            }
        }

        [TestMethod]
        public void CalculatedColumnFormula_RemoveFormulas()
        {
            using (var p = OpenPackage("CalculatedColumnFormulaRemove1.xlsx", true))
            {
                // Set up a worksheet containing a table
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";
                ws.Cells["C1"].Value = "Col3";
                ws.Cells["A2"].Value = 1;
                ws.Cells["B2"].Value = 2;
                var table1 = ws.Tables.Add(ws.Cells["A1:C2"], "Table1");
                var formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
                table1.Columns[2].CalculatedColumnFormula = formula;

                // Check the calculated column formula
                Assert.AreEqual(formula, ws.Cells["C2"].Formula);
                Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

                // Remove all formulas from the table
                table1.Range.ClearFormulas();
                table1.Range.ClearFormulaValues();

                // Check the calculated column formula is no longer there
                Assert.IsTrue(string.IsNullOrEmpty(table1.Columns["Col3"].CalculatedColumnFormula));
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void CalculatedColumnFormula_RemoveFormulas_AddRow()
        {
            using (var p = OpenPackage("CalculatedColumnFormulaRemove2.xlsx", true))
            {
                // Set up a worksheet containing a table
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";
                ws.Cells["C1"].Value = "Col3";
                ws.Cells["A2"].Value = 1;
                ws.Cells["B2"].Value = 2;
                var table1 = ws.Tables.Add(ws.Cells["A1:C2"], "Table1");
                var formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
                table1.Columns[2].CalculatedColumnFormula = formula;

                // Check the calculated column formula
                Assert.AreEqual(formula, ws.Cells["C2"].Formula);
                Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

                // Remove all formulas from the table
                table1.Range.ClearFormulas();
                table1.Range.ClearFormulaValues();
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["C2"].Formula));

                // Add a row to the table
                table1.InsertRow(1);

                // Check the formula has not been reinserted
                Assert.IsTrue(string.IsNullOrEmpty(ws.Cells["C2"].Formula));
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void CalculatedColumnFormula_OneCellDifferent_AddRow()
        {
            using (var p = OpenPackage("CalculatedColumnFormulaRemove3.xlsx", true))
            {
                // Set up a worksheet containing a table
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";
                ws.Cells["C1"].Value = "Col3";
                ws.Cells["A2"].Value = 1;
                ws.Cells["B2"].Value = 2;
                ws.Cells["A3"].Value = 3;
                ws.Cells["B3"].Value = 4;
                ws.Cells["A4"].Value = 5;
                ws.Cells["B4"].Value = 6;
                var table1 = ws.Tables.Add(ws.Cells["A1:C4"], "Table1");
                var formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
                table1.Columns[2].CalculatedColumnFormula = formula;

                // Check the calculated column formula has been added to each cell
                Assert.AreEqual(formula, ws.Cells["C2"].Formula);
                Assert.AreEqual(formula, ws.Cells["C3"].Formula);
                Assert.AreEqual(formula, ws.Cells["C4"].Formula);
                Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

                // Remove the calculated column formula from one row and use a different formula instead
                ws.Cells["C3"].ClearFormulas();
                ws.Cells["C3"].ClearFormulaValues();
                var differentFormula = "Table1[[#This Row],[Col1]]";
                ws.Cells["C3"].Formula = differentFormula;
                Assert.AreEqual(differentFormula, ws.Cells["C3"].Formula);

                // Add a new row to the bottom of the table
                table1.AddRow();

                // Check that the new row has the formula
                Assert.AreEqual(formula, ws.Cells["C5"].Formula);
                Assert.AreEqual(formula, table1.Columns["Col3"].CalculatedColumnFormula);

                // Check the cell where we used a different formula hasn't changed
                Assert.AreEqual(differentFormula, ws.Cells["C3"].Formula);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void CreateTableAfterDeletingAMergedCell()
        {
            // Reproduce issue 780
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Prepare some data
                worksheet.Cells["A1"].Value = "Column 1";
                worksheet.Cells["A2"].Value = 1;
                worksheet.Cells["B1"].Value = "Column 2";
                worksheet.Cells["B2"].Value = 2;

                // Merge cells in row 4 (not related to the data above)
                worksheet.Cells["A4:B4"].Merge = true;
                // Delete the row that has the merged cells
                worksheet.DeleteRow(4);

                // Create a table
                var tableCells = worksheet.Cells["A1:B2"];
                var table = worksheet.Tables.Add(tableCells, "table"); // --> This triggers a NullReferenceException
            }
        }

        [TestMethod]
        public void ShowTotalWhenValueBelowIt()
        {
            using (var package = OpenPackage("ShowTotalInsert.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Tables");

                //Cause of issue. (Specifically sheet.cells[11,1]
                for (int i = 1; i < 25; i++)
                {
                    sheet.Cells[1 + i, 1].Formula = $"\"Number:{i}\"";
                }

                sheet.Cells["A1"].Value = "Month";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["C1"].Value = "VAT";
                sheet.Cells["D1"].Value = "Total";

                var table = sheet.Tables.Add(new ExcelAddress("A1:E10"), "testTable");
                table.ShowHeader = true;
                table.ShowFirstColumn = true;
                table.TableStyle = TableStyles.Dark2;

                sheet.Cells["A1"].Value = "testStuff";

                sheet.Cells["C3:C5"].Value = 3;
                sheet.Cells["D3:E3"].Value = 4;
                sheet.Cells["E10"].Value = 5;

                sheet.Cells["F11"].Value = "Don't clear me";
                var noHeader = package.Workbook.Worksheets.Add("noHeader", sheet);
                var stackTest = package.Workbook.Worksheets.Add("stackTest", sheet);

                table.ShowTotal = true;
                table.Columns[2].TotalsRowFunction = RowFunctions.Sum;

                sheet.Calculate();

                Assert.AreEqual(sheet.Cells["A10"].Value, "Number:9");
                Assert.AreEqual(sheet.Cells["A11"].Value, null);
                Assert.AreEqual(sheet.Cells["A12"].Value, "Number:11");
                Assert.AreEqual(sheet.Cells["F11"].Value, "Don't clear me");

                noHeader.Tables[0].ShowHeader = false;
                noHeader.Tables[0].ShowTotal = true;
                noHeader.Tables[0].Columns[2].TotalsRowFunction = RowFunctions.Sum;

                noHeader.Calculate();

                Assert.AreEqual(noHeader.Cells["A10"].Value, "Number:9");
                Assert.AreEqual(noHeader.Cells["A11"].Value, null);
                Assert.AreEqual(noHeader.Cells["A12"].Value, "Number:11");
                Assert.AreEqual(noHeader.Cells["F11"].Value, "Don't clear me");

                stackTest.Tables[0].ShowTotal = true;
                stackTest.Tables[0].ShowTotal = false;
                stackTest.Tables[0].ShowTotal = true;
                stackTest.Tables[0].ShowTotal = false;
                stackTest.Tables[0].ShowTotal = true;
                stackTest.Tables[0].ShowTotal = false;
                stackTest.Tables[0].ShowTotal = true;
                stackTest.Tables[0].ShowTotal = false;

                stackTest.Calculate();

                Assert.AreEqual(stackTest.Cells["A10"].Value, "Number:9");
                Assert.AreEqual(stackTest.Cells["A11"].Value, null);
                Assert.AreEqual(stackTest.Cells["A12"].Value, "Number:11");
                Assert.AreEqual(stackTest.Cells["F11"].Value, "Don't clear me");

                SaveAndCleanup(package);
            }
        }
        private void InitDataTable(out DataTable table)
        {
            table = new DataTable();
            table.Columns.Add("Country", typeof(string));
            table.Columns.Add("Population", typeof(int));
            var areaCol = table.Columns.Add("Area", typeof(int));
            areaCol.Caption = "Area (km2)";

            table.Rows.Add("Sweden", 10409248, 450295);
            table.Rows.Add("Norway", 5402171, 385178);
            table.Rows.Add("Netherlands", 17553530, 41198);
        }

        [TestMethod]
        public void ColNamesAndColumnNameShouldBeEqual()
        {
            using (var p = OpenPackage("totalTable.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("colWs");

                DataTable data = new DataTable();
                InitDataTable(out data);

                TableStyles style = TableStyles.Dark1;
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(data, true, style);

                // configure the table
                var table = sheet.Tables.GetFromRange(tableRange);
                table.Sort(x => x.SortBy.ColumnNamed("Population", eSortOrder.Descending));
                table.ShowTotal = true;
                table.Columns[0].TotalsRowLabel = "Total";
                table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[2].TotalsRowFunction = RowFunctions.Sum;

                // add column for population density
                table.Columns.Add(1);
                tableRange = table.Range;
                table.Columns[3].CalculatedColumnFormula = $"{table.Name}[[#This Row],[Population]]/{table.Name}[[#This Row],[Area (km2)]]";
                table.Columns[3].Name = "Density";
                table.Columns[3].TotalsRowFunction = RowFunctions.Average;
                sheet.Calculate();

                var totalRow = tableRange.End.Row;
                Assert.AreEqual(table.Columns.GetIndexOfColName("Density"), 3);
                Assert.AreEqual(154.40629115594086d, sheet.Cells[totalRow, 4].Value);
            }

        }

        [TestMethod]
        public void CalculatedColumnFormula_SetToEmptyString_CellStyle()
        {
            using (var pck = new ExcelPackage())
            {
                // Set up a worksheet containing a table
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                wks.Cells["A1"].Value = "Col1";
                wks.Cells["B1"].Value = "Col2";
                wks.Cells["C1"].Value = "Col3";
                wks.Cells["A2"].Value = 1;
                wks.Cells["B2"].Value = 2;
                var table1 = wks.Tables.Add(wks.Cells["A1:C5"], "Table1");
                var formula = "Table1[[#This Row],[Col1]]+Table1[[#This Row],[Col2]]";
                table1.Columns[2].CalculatedColumnFormula = formula;
                wks.Calculate();

                // Add a style to the cell containing the formula
                wks.Cells["B2:C3"].Style.Font.Bold = true;
                wks.Cells["B2:C3"].Style.Font.Size = 16;
                wks.Cells["B2:C3"].Style.Font.Color.SetColor(eThemeSchemeColor.Text2);
                wks.Cells["B2:C3"].Style.Font.Color.Tint = 0.39997558519241921m;

                // Check the style has been applied
                Assert.AreEqual(true, wks.Cells["B3"].Style.Font.Bold);
                Assert.AreEqual(16, wks.Cells["B3"].Style.Font.Size, 1E-3);
                Assert.AreEqual(eThemeSchemeColor.Text2, wks.Cells["B3"].Style.Font.Color.Theme);
                Assert.AreEqual(0.39997558519241921m, wks.Cells["B3"].Style.Font.Color.Tint);

                // Remove the calculated column formula from the table
                table1.Columns["Col3"].CalculatedColumnFormula = "";

                // Check the style hasn't changed
                Assert.AreEqual(true, wks.Cells["B3"].Style.Font.Bold);
                Assert.AreEqual(16, wks.Cells["B3"].Style.Font.Size, 1E-3);
                Assert.AreEqual(eThemeSchemeColor.Text2, wks.Cells["B3"].Style.Font.Color.Theme);
                Assert.AreEqual(0.39997558519241921m, wks.Cells["B3"].Style.Font.Color.Tint);
            }
        }
        [TestMethod]
        public void TestOverWriteColumnNamesWithCells()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:D5");
                ws.SetValue("A2", "Column3");
                ws.SetValue("B2", "Column4");

                var table = ws.Tables.Add(range, "newTable");

                ws.SetValue("A1", "Items");
                ws.SetValue("B1", "Years");

                Assert.AreEqual("Column1", table.Columns[0].Name);
                Assert.AreEqual("Column2", table.Columns[1].Name);
                Assert.AreEqual("Column3", table.Columns[2].Name);
                Assert.AreEqual("Column4", table.Columns[3].Name);

                table.SyncColumnNames(ApplyDataFrom.CellsToColumnNames, false);

                Assert.AreEqual("Items", table.Columns[0].Name);
                Assert.AreEqual("Years", table.Columns[1].Name);
                Assert.AreEqual("Column3", table.Columns[2].Name);
                Assert.AreEqual("Column4", table.Columns[3].Name);

                Assert.AreEqual("Items", ws.Cells["A1"].Value);
                Assert.AreEqual("Years", ws.Cells["B1"].Value);
                Assert.AreEqual(null, ws.Cells["C1"].Value);
                Assert.AreEqual(null, ws.Cells["D1"].Value);

                table.SyncColumnNames(ApplyDataFrom.CellsToColumnNames);

                Assert.AreEqual("Column3", ws.Cells["C1"].Value);
                Assert.AreEqual("Column4", ws.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void RemovingAndAddingHeaders()
        {
            using (var package = OpenPackage("RemovingAddingHeaders.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:C5");
                ws.SetValue("A2", "Column3");
                ws.SetValue("B2", "Column4");

                var table = ws.Tables.Add(range, "newTable");

                table.ShowHeader = false;

                Assert.AreEqual(null, ws.Cells["A1"].Value);
                Assert.AreEqual(null, ws.Cells["B1"].Value);

                table.ShowHeader = true;

                Assert.AreEqual(null, ws.Cells["A1"].Value);
                Assert.AreEqual(null, ws.Cells["B1"].Value);

                table.SyncColumnNames(ApplyDataFrom.ColumnNamesToCells);

                Assert.AreEqual("Column1", ws.Cells["A1"].Value);
                Assert.AreEqual("Column2", ws.Cells["B1"].Value);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreatingTableWithColumnNamesAlready()
        {
            using (var package = OpenPackage("DuplicateColumnNames.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:C5");
                ws.SetValue("A1", "Column3");
                ws.SetValue("B1", "Column4");

                var table = ws.Tables.Add(range, "newTable");

                table.ShowHeader = true;

                Assert.AreEqual("Column3", ws.Cells["A1"].Value);
                Assert.AreEqual("Column4", ws.Cells["B1"].Value);
                Assert.AreEqual(null, ws.Cells["C1"].Value);
                Assert.AreEqual("Column32", table.Columns[2].Name);
                table.Columns.Add(1);
                Assert.AreEqual("Column42", table.Columns[3].Name);
                Assert.AreEqual("Column42", ws.Cells["D1"].Value);


                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreatingAndSavingTableSameColumnName()
        {
            using (var package = OpenPackage("TableSameColName.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:B5");
                ws.SetValue("A1", "AColumn");
                ws.SetValue("B1", "AColumn");

                var table = ws.Tables.Add(range, "newTable");

                table.ShowHeader = true;

                Assert.AreEqual("AColumn", table.Columns[0].Name);
                Assert.AreEqual("Column2", table.Columns[1].Name);

                Assert.AreEqual("AColumn", ws.Cells["A1"].Value);
                Assert.AreEqual("AColumn", ws.Cells["B1"].Value);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreatingAndSavingEmptyColNames()
        {
            using (var package = OpenPackage("TableEmptyColNames.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:B5");
                ws.SetValue("A2", "something");
                ws.SetValue("B2", "somethingElse");

                var table = ws.Tables.Add(range, "newTable");

                Assert.AreEqual("Column1", table.Columns[0].Name);
                Assert.AreEqual("Column2", table.Columns[1].Name);

                Assert.AreEqual(null, ws.Cells["A1"].Value);
                Assert.AreEqual(null, ws.Cells["B1"].Value);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void AddingColumnWhenBaseColumnNameExists()
        {
            using (var package = OpenPackage("TableSameColNameAdding.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:C5");
                ws.SetValue("A1", "AColumn");
                ws.SetValue("B1", "AnotherColumn");
                ws.SetValue("C1", "ThirdColumn");

                var table = ws.Tables.Add(range, "newTable");

                table.ShowHeader = true;

                table.Columns.Add(1);

                Assert.AreEqual("Column4", ws.Cells["D1"].Value);
                Assert.AreEqual("ThirdColumn", table.Columns[2].Name);

                table.Columns[2].Name = "Column5";

                table.Columns.Add(1);
                Assert.AreEqual("Column52",table.Columns[4].Name);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ThrowsWhenAttemptingToAddSameName()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("TESTTABLE");
                var range = new ExcelAddress("A1:C5");
                ws.SetValue("A1", "AColumn");
                ws.SetValue("B1", "AnotherColumn");
                ws.SetValue("C1", "ThirdColumn");

                var table = ws.Tables.Add(range, "newTable");

                table.Columns[1].Name = "AColumn";
            }
        }
    }
}
