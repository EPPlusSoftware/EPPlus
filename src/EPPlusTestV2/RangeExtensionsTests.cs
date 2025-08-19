using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class RangeExtensionsTests
    {
        [TestMethod]
        public void ShouldSkipColumns()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;
                
                var range = sheet.Cells["A1:B3"].SkipColumns(1);
                Assert.AreEqual("B1:B3", range.Address);

                range = sheet.Cells["A1:C3"].SkipColumns(1);
                Assert.AreEqual("B1:C3", range.Address);
                
                range = sheet.Cells["A1:C3"].SkipColumns(2);
                Assert.AreEqual("C1:C3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldSkipColumns2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                var range = sheet.Cells["B2:D10"].SkipColumns(1);
                Assert.AreEqual("C2:D10", range.Address);
            }
        }

        [TestMethod, ExpectedException(typeof(IndexOutOfRangeException))]
        public void SkipColumnsShouldThrowIfNbrOfColumnsIsTooLarge()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                var range = sheet.Cells["A1:B3"].SkipColumns(4);
            }
        }

        [TestMethod]
        public void ShouldSkipRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:B3"].SkipRows(1);
                Assert.AreEqual("A2:B3", range.Address);

                range = sheet.Cells["A1:C3"].SkipRows(1);
                Assert.AreEqual("A2:C3", range.Address);

                range = sheet.Cells["A1:C3"].SkipRows(2);
                Assert.AreEqual("A3:C3", range.Address);

                range = sheet.Cells["A2:B14"].SkipRows(1);
                Assert.AreEqual("A3:B14", range.Address);
            }
        }

        [TestMethod]
        public void ShouldSkipRows2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                var range = sheet.Cells["A2:B14"].SkipRows(1);
                Assert.AreEqual("A3:B14", range.Address);
            }
        }

        [TestMethod, ExpectedException(typeof(IndexOutOfRangeException))]
        public void SkipRowsShouldThrowIfNbrOfColumnsIsTooLarge()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                var range = sheet.Cells["A1:B3"].SkipRows(4);
            }
        }

        [TestMethod]
        public void ShouldTakeColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:B3"].TakeColumns(1);
                Assert.AreEqual("A1:A3", range.Address);

                range = sheet.Cells["A1:C3"].TakeColumns(2);
                Assert.AreEqual("A1:B3", range.Address);

                range = sheet.Cells["A1:C3"].TakeColumns(3);
                Assert.AreEqual("A1:C3", range.Address);

                range = sheet.Cells["A1:C3"].TakeColumns(5);
                Assert.AreEqual("A1:C3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeSingleColumn()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:C3"].TakeSingleColumn(0);
                Assert.AreEqual("A1:A3", range.Address);

                range = sheet.Cells["A1:C3"].TakeSingleColumn(1);
                Assert.AreEqual("B1:B3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:B3"].TakeRows(1);
                Assert.AreEqual("A1:B1", range.Address);

                range = sheet.Cells["A1:C3"].TakeRows(2);
                Assert.AreEqual("A1:C2", range.Address);

                range = sheet.Cells["A1:C3"].TakeRows(3);
                Assert.AreEqual("A1:C3", range.Address);

                range = sheet.Cells["A1:C3"].TakeRows(5);
                Assert.AreEqual("A1:C3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeSingleRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:C3"].TakeSingleRow(0);
                Assert.AreEqual("A1:C1", range.Address);

                range = sheet.Cells["A1:C3"].TakeSingleRow(1);
                Assert.AreEqual("A2:C2", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeColumnsBetween()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:C3"].TakeColumnsBetween(0, 1);
                Assert.AreEqual("A1:A3", range.Address);

                range = sheet.Cells["A1:C3"].TakeColumnsBetween(1, 2);
                Assert.AreEqual("B1:C3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeColumnsBetween2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                var range = sheet.Cells["B1:D3"].TakeColumnsBetween(0, 1);
                Assert.AreEqual("B1:B3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeRowsBetween()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var range = sheet.Cells["A1:C3"].TakeRowsBetween(0, 1);
                Assert.AreEqual("A1:C1", range.Address);

                range = sheet.Cells["A1:C3"].TakeRowsBetween(1, 2);
                Assert.AreEqual("A2:C3", range.Address);
            }
        }

        [TestMethod]
        public void ShouldTakeSingleCell()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C1"].Value = 7;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 9;

                var cell = sheet.Cells["A1:C3"].TakeSingleCell(1, 1);
                Assert.AreEqual(5, cell.Value);
            }
        }
    }
}
