using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
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
        }
    }
}
