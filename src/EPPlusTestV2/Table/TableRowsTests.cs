using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.Worksheet.XmlWriter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Text;

namespace EPPlusTest.Table
{
    [TestClass]
    public class TableRowsTests : TestBase
    {
        [TestMethod]
        public void TestNormalTable_ColumnNames()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            sheet.Cells["C1"].Value = "Col3";
            for (var col = 1; col < 4; col++)
            {
                for (var row = 2; row < 4; row++)
                {
                    sheet.Cells[row, col].Value = col * row;
                }
            }

            var table = sheet.Tables.Add(sheet.Cells["A1:C3"], "Table1");
            Assert.AreEqual(2, table.DataRows.Count());
            Assert.AreEqual("A2:C2", table.DataRows[0].RowRange.Address);
            Assert.IsFalse(table.DataRows[0].IsHidden);
            Assert.AreEqual(2, table.DataRows[0].GetValue<int>("Col1"));
            Assert.AreEqual(4, table.DataRows[0].GetValue<int>("Col2"));
            Assert.AreEqual(6, table.DataRows[0].GetValue<int>("Col3"));
            Assert.AreEqual("A3:C3", table.DataRows[1].RowRange.Address);
            Assert.AreEqual(3, table.DataRows[1].GetValue<int>("Col1"));
            Assert.AreEqual(6, table.DataRows[1].GetValue<int>("Col2"));
            Assert.AreEqual(9, table.DataRows[1].GetValue<int>("Col3"));
        }

        [TestMethod]
        public void TestNormalTable_ColumnIndex()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            sheet.Cells["C1"].Value = "Col3";
            for (var col = 1; col < 4; col++)
            {
                for (var row = 2; row < 4; row++)
                {
                    sheet.Cells[row, col].Value = col * row;
                }
            }

            var table = sheet.Tables.Add(sheet.Cells["A1:C3"], "Table1");
            Assert.AreEqual(2, table.DataRows.Count());
            Assert.AreEqual("A2:C2", table.DataRows[0].RowRange.Address);
            Assert.IsFalse(table.DataRows[0].IsHidden);
            Assert.AreEqual(2, table.DataRows[0].GetValue<int>(0));
            Assert.AreEqual(4, table.DataRows[0].GetValue<int>(1));
            Assert.AreEqual(6, table.DataRows[0].GetValue<int>(2));
            Assert.AreEqual("A3:C3", table.DataRows[1].RowRange.Address);
            Assert.AreEqual(3, table.DataRows[1].GetValue<int>(0));
            Assert.AreEqual(6, table.DataRows[1].GetValue<int>(1));
            Assert.AreEqual(9, table.DataRows[1].GetValue<int>(2));
        }

        [TestMethod]
        public void TestTableWithAutofilter()
        {
            using var p = OpenTemplatePackage("TableRowsAutofilter.xlsx");
            var tbl = p.Workbook.Worksheets[0].Tables[0];
            Assert.AreEqual(3, tbl.DataRows.Count());
            Assert.IsTrue(tbl.DataRows[0].IsHidden);
            Assert.IsFalse(tbl.DataRows[1].IsHidden);
            Assert.IsFalse(tbl.DataRows[2].IsHidden);
            Assert.AreEqual(4, tbl.DataRows[1].GetValue<int>("Col1"));
            Assert.AreEqual("5", tbl.DataRows[1].GetValue<string>("Col2"));

            var visibleRows = tbl.DataRows.Where(r => r.IsHidden == false);
            Assert.AreEqual(2, visibleRows.Count());
        }

        [TestMethod]
        public void TestTableWithAutofilter_AddNewRow()
        {
            using var p = OpenTemplatePackage("TableRowsAutofilter.xlsx");
            var tbl = p.Workbook.Worksheets[0].Tables[0];
            Assert.AreEqual(3, tbl.DataRows.Count());
            var emptyRows = tbl.DataRows.Where(r => r.IsEmpty);
            Assert.AreEqual(0, emptyRows.Count());
            var newRow = tbl.DataRows.AddNewRow();
            Assert.AreEqual(4, tbl.DataRows.Count());
            emptyRows = tbl.DataRows.Where(r => r.IsEmpty);
            Assert.AreEqual(1, emptyRows.Count());

            Assert.IsTrue(newRow.IsEmpty);
            newRow.SetValue("Col1", 10);
            Assert.IsFalse(newRow.IsEmpty);
            Assert.AreEqual(10, p.Workbook.Worksheets[0].Cells["C7"].Value, "C7 was not 10");
        }

        [TestMethod]
        public void TestTableWithAutofilter_AddNewRows()
        {
            using var p = OpenTemplatePackage("TableRowsAutofilter.xlsx");
            var tbl = p.Workbook.Worksheets[0].Tables[0];
            Assert.AreEqual(3, tbl.DataRows.Count());
            var emptyRows = tbl.DataRows.Where(r => r.IsEmpty);
            Assert.AreEqual(0, emptyRows.Count());
            var newRows = tbl.DataRows.AddNewRows(2);
            Assert.AreEqual(5, tbl.DataRows.Count());
            emptyRows = tbl.DataRows.Where(r => r.IsEmpty);
            Assert.AreEqual(2, emptyRows.Count());

            Assert.IsTrue(newRows.First().IsEmpty);
            newRows.First().SetValue("Col1", 10);
            newRows.Last().SetValue("Col1", 11);
            Assert.IsFalse(newRows.First().IsEmpty);
            Assert.IsFalse(newRows.Last().IsEmpty);
            Assert.AreEqual(10, p.Workbook.Worksheets[0].Cells["C7"].Value, "C7 was not 10");
            Assert.AreEqual(11, p.Workbook.Worksheets[0].Cells["C8"].Value, "C8 was not 11");
        }

        [TestMethod]
        public void InsertNewRow()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            tbl.DataRows[0].SetValue("Col1", 1).SetValue("Col2", 2);
            var newRows = tbl.DataRows.AddNewRows(4);
            var n = 3;
            foreach(var newRow in newRows)
            {
                newRow.SetValue("Col1", n++);
                newRow.SetValue("Col2", n++);
            }

            Assert.AreEqual(5, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(7, tbl.DataRows[3].GetValue("Col1"));

            var insertedRow = tbl.DataRows.InsertNewRow(2);
            insertedRow.SetValue("Col1", 100);

            Assert.AreEqual(6, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(100, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[3].GetValue("Col1"));
        }

        [TestMethod]
        public void InsertNewRows()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            tbl.DataRows[0].SetValue("Col1", 1).SetValue("Col2", 2);
            var newRows = tbl.DataRows.AddNewRows(4);
            var n = 3;
            foreach (var newRow in newRows)
            {
                newRow.SetValue("Col1", n++);
                newRow.SetValue("Col2", n++);
            }

            Assert.AreEqual(5, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(7, tbl.DataRows[3].GetValue("Col1"));

            var insertedRows = tbl.DataRows.InsertNewRows(2, 2);
            insertedRows.ElementAt(0).SetValue("Col1", 100);
            insertedRows.ElementAt(1).SetValue("Col1", 101);

            Assert.AreEqual(7, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(100, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(101, tbl.DataRows[3].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[4].GetValue("Col1"));
            Assert.AreEqual(7, tbl.DataRows[5].GetValue("Col1"));
        }

        [TestMethod]
        public void DeleteRows()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            tbl.DataRows[0].SetValue("Col1", 1).SetValue("Col2", 2);
            var newRows = tbl.DataRows.AddNewRows(4);
            var n = 3;
            foreach (var newRow in newRows)
            {
                newRow.SetValue("Col1", n++);
                newRow.SetValue("Col2", n++);
            }

            Assert.AreEqual(5, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(7, tbl.DataRows[3].GetValue("Col1"));

            tbl.DataRows.DeleteRows(2, 2);

            Assert.AreEqual(3, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(9, tbl.DataRows[2].GetValue("Col1"));
        }

        [TestMethod]
        public void DeleteMultipleRows()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            tbl.DataRows[0].SetValue("Col1", 1).SetValue("Col2", 2);
            var newRows = tbl.DataRows.AddNewRows(4);
            var n = 3;
            foreach (var newRow in newRows)
            {
                newRow.SetValue("Col1", n++);
                newRow.SetValue("Col2", n++);
            }

            Assert.AreEqual(5, tbl.DataRows.Count());
            var lastRow = tbl.DataRows.Last();
            Assert.AreEqual(10, lastRow.GetValue<int>("Col2"));
            Assert.AreEqual(4, lastRow.RowIx);
            var row1 = tbl.DataRows[2];
            Assert.AreEqual(5, row1.GetValue<int>(0));
            var row2 = tbl.DataRows[3];
            row1.Delete();
            Assert.IsTrue(row1.IsDeleted);
            Assert.IsFalse(row2.IsDeleted);
            Assert.AreEqual(3, lastRow.RowIx);
            row2.Delete();
            Assert.IsTrue(row2.IsDeleted);
            Assert.AreEqual(3, tbl.DataRows.Count());
            Assert.AreEqual(2, lastRow.RowIx);
            Assert.AreEqual(10, lastRow.GetValue<int>("Col2"));
        }

        [TestMethod]
        public void ClearTable()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            tbl.DataRows[0].SetValue("Col1", 1).SetValue("Col2", 2);
            var newRows = tbl.DataRows.AddNewRows(4);
            var n = 3;
            foreach (var newRow in newRows)
            {
                newRow.SetValue("Col1", n++);
                newRow.SetValue("Col2", n++);
            }

            Assert.AreEqual(5, tbl.DataRows.Count());
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(5, tbl.DataRows[2].GetValue("Col1"));
            Assert.AreEqual(7, tbl.DataRows[3].GetValue("Col1"));

            tbl.DataRows.Clear();

            Assert.AreEqual(1, tbl.DataRows.Count());
            Assert.IsTrue(tbl.DataRows[0].IsEmpty);
        }

        [TestMethod]
        public void SetValues1()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            Assert.AreEqual(1, tbl.DataRows.Count());
            Assert.IsTrue(tbl.DataRows[0].IsEmpty);
            tbl.DataRows[0].SetValues(1, 2);
            Assert.IsFalse(tbl.DataRows[0].IsEmpty);
            Assert.AreEqual(1, tbl.DataRows[0].GetValue("Col1"));
            Assert.AreEqual(2, tbl.DataRows[0].GetValue("Col2"));
            var newRow = tbl.DataRows.AddNewRow();
            newRow.SetValues(new int[] { 3, 4 });
            Assert.IsFalse(tbl.DataRows[1].IsEmpty);
            Assert.AreEqual(3, tbl.DataRows[1].GetValue("Col1"));
            Assert.AreEqual(4, tbl.DataRows[1].GetValue("Col2"));
        }

        [TestMethod, ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void SetValues2()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Col1";
            sheet.Cells["B1"].Value = "Col2";
            var tbl = sheet.Tables.Add(sheet.Cells["A1:B2"], "Table1");
            var newRow = tbl.DataRows.AddNewRow();
            newRow.SetValues(new int[] { 3, 4, 5 });
        }
        [TestMethod]
        public void EnsureDataRowsWorksWithCustomColumnNames()
        {
            using (var p = OpenPackage("tablePackage.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("AWs");

                var table = ws.Tables.Add(ws.Cells["A1:D10"], "SomeTable");
                table.ShowHeader = true;

                for (int i = 1; i < 5; i++)
                {
                    ws.Cells[1, i].Value = $"CustomColumn{i}";
                }

                ws.Cells["A2:D10"].Formula = "COLUMN() + ROW()";

                ws.Calculate();

                table.DataRows[0].SetValue("Column4", 75);

                table.SyncColumnNames(OfficeOpenXml.Table.ApplyDataFrom.CellsToColumnNames);

                foreach (var col in table.Columns.Where(col => col.Name == "CustomColumn1" || col.Name == "CustomColumn3"))
                {
                    foreach (var row in table.DataRows)
                    {
                        table.DataRows[0].SetValue(col.Name, 367);

                        table.DataRows[3].SetValue(col.Name, 333);
                    }
                }

                //Ensure data rows are accurate
                Assert.AreEqual(367, ws.Cells["A2"].Value);
                Assert.AreEqual(367, ws.Cells["C2"].Value);
                Assert.AreEqual(333, ws.Cells["A5"].Value);
                Assert.AreEqual(333, ws.Cells["C5"].Value);
                Assert.AreEqual(75, ws.Cells["D2"].Value);

                //Ensure calculated values between set values have not been set
                Assert.AreEqual(4d, ws.Cells["B2"].Value);
                Assert.AreEqual(7d, ws.Cells["B5"].Value);
                Assert.AreEqual(7d, ws.Cells["D3"].Value);
                Assert.AreEqual(9d, ws.Cells["D5"].Value);

                SaveAndCleanup(p);
            }
        }
    }
}
