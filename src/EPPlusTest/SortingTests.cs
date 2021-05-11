using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class SortingTests
    {
        [TestMethod]
        public void ShouldSetSortState()
        {
            using(var package = new ExcelPackage())
            {
                var rnd = new Random();
                var sheet = package.Workbook.Worksheets.Add("Test");
                Assert.IsNull(sheet.SortState, "Worksheet.SortState was not null");
                for (var x = 1; x < 10; x++)
                {
                    sheet.Cells[x, 1].Value = rnd.Next(1, 20);
                }
                sheet.Cells["A1:A10"].Sort(0, true);
                Assert.IsNotNull(sheet.SortState, "Worksheet.SortState was null");
                Assert.AreEqual(1, sheet.SortState.SortConditions.Count());
            }
        }

        [TestMethod]
        public void ShouldSortMultipleColumns()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[2, 1].Value = 6;
                sheet.Cells[3, 1].Value = 2;
                sheet.Cells[4, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = 2;
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[4, 2].Value = 4;
                sheet.Cells[1, 3].Value = "A";
                sheet.Cells[2, 3].Value = "D";
                sheet.Cells[3, 3].Value = "C";
                sheet.Cells[4, 3].Value = "B";
                sheet.Cells["A1:C4"].Sort(new int[2] { 0, 2 }, new bool[2] { false, true });
                Assert.IsNotNull(sheet.SortState);
                Assert.AreEqual(2, sheet.SortState.SortConditions.Count());
            }
        }

        [TestMethod]
        public void ShouldSortSingleColumn()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells[1, 1].Value = 4;
                sheet.Cells[2, 1].Value = 3;
                sheet.Cells[3, 1].Value = 2;

                // sort ascending
                sheet.Cells["A1:A3"].Sort();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
                Assert.AreEqual(3, sheet.Cells["A2"].Value);
                Assert.AreEqual(4, sheet.Cells["A3"].Value);

                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 2;

                // sort descending
                sheet.Cells["A1:A3"].Sort(0, true);
                Assert.AreEqual(4, sheet.Cells["A1"].Value);
                Assert.AreEqual(3, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void ShouldConfigureSortOptionsForSingleColumn()
        {
            var options = new RangeSortOptions();
            options.SortBy.Column(0);
            Assert.AreEqual(1, options.ColumnIndexes.Count);
            Assert.AreEqual(1, options.Descending.Count);
            Assert.AreEqual(0, options.ColumnIndexes.First());
            Assert.IsFalse(options.Descending.First());
        }

        [TestMethod]
        public void ShouldConfigureSortOptionsForMultiColumns()
        {
            var options = new RangeSortOptions();
            options
                .SortBy.Column(0)
                .ThenSortBy.Column(3, eSortDirection.Descending)
                .ThenSortBy.Column(2);
            Assert.AreEqual(3, options.ColumnIndexes.Count);
            Assert.AreEqual(3, options.Descending.Count);
            Assert.AreEqual(0, options.ColumnIndexes.First());
            Assert.IsFalse(options.Descending.First());
            Assert.AreEqual(3, options.ColumnIndexes[1]);
            Assert.IsTrue(options.Descending[1]);
        }

        [TestMethod]
        public void ShouldSortMultipleColumnsWithOptions()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[2, 1].Value = 6;
                sheet.Cells[3, 1].Value = 2;
                sheet.Cells[4, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = 2;
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[4, 2].Value = 4;
                sheet.Cells[1, 3].Value = "A";
                sheet.Cells[2, 3].Value = "D";
                sheet.Cells[3, 3].Value = "C";
                sheet.Cells[4, 3].Value = "B";

                sheet.Cells["A1:C4"].Sort(options => options
                                                        .SortBy.Column(0)
                                                        .ThenSortBy.Column(2, eSortDirection.Descending));
                Assert.IsNotNull(sheet.SortState);
                Assert.AreEqual(2, sheet.SortState.SortConditions.Count());
                Assert.IsTrue(sheet.SortState.SortConditions.Last().Descending);
                Assert.AreEqual(6, sheet.Cells[4, 1].Value);
                Assert.AreEqual("B", sheet.Cells[1, 3].Value);
            }
        }

        [TestMethod]
        public void ShouldSortByCustomList()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells[1, 1].Value = "Blue";
                sheet.Cells[2, 1].Value = "Red";
                sheet.Cells[3, 1].Value = "Yellow";
                sheet.Cells[4, 1].Value = "Blue";

                sheet.Cells[1, 2].Value = 2;
                sheet.Cells[2, 2].Value = 1;
                sheet.Cells[3, 2].Value = 1;
                sheet.Cells[4, 2].Value = 1;

                sheet.Cells["A1:B4"].Sort(x => x.SortBy.Column(0).UsingCustomList("Red", "Yellow", "Blue").ThenSortBy.Column(1));

                Assert.AreEqual("Red", sheet.Cells[1, 1].Value);
                Assert.AreEqual("Yellow", sheet.Cells[2, 1].Value);
                Assert.AreEqual("Blue", sheet.Cells[3, 1].Value);
                Assert.AreEqual(1, sheet.Cells[3, 2].Value);
                Assert.AreEqual("Blue", sheet.Cells[4, 1].Value);
                Assert.AreEqual(2, sheet.Cells[4, 2].Value);
            }
        }

        private ExcelTable CreateTable(ExcelWorksheet sheet, bool addTotalsRow = true)
        {
            // header
            sheet.Cells[1, 1].Value = "Header1";
            sheet.Cells[1, 2].Value = "Header2";
            sheet.Cells[1, 3].Value = "Header3";
            // row 1
            sheet.Cells[2, 1].Value = 10;
            sheet.Cells[2, 2].Value = 2;
            sheet.Cells[2, 3].Value = 3;
            // row 2
            sheet.Cells[3, 1].Value = 5;
            sheet.Cells[3, 2].Value = 2;
            sheet.Cells[3, 3].Value = 3;

            var table = sheet.Tables.Add(sheet.Cells["A1:C3"], "myTable");
            table.TableStyle = TableStyles.Dark1;
            table.ShowTotal = addTotalsRow;
            table.Columns[0].TotalsRowFunction = RowFunctions.Sum;
            table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
            return table;
        }

        [TestMethod]
        public void TableSortByColumnIndex()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var table = CreateTable(sheet);

                table.Sort(x => x.SortBy.Column(0));

                Assert.AreEqual(5, sheet.Cells[2, 1].Value);
                Assert.AreEqual(10, sheet.Cells[3, 1].Value);
                Assert.IsNotNull(table.SortState, "SortState was null");
                Assert.IsNotNull(table.SortState.SortConditions, "SortState.SortConditions was null");
                Assert.IsFalse(table.SortState.SortConditions.First().Descending, "First SortCondition was not descending");
            }
        }

        [TestMethod]
        public void TableSortByColumnName()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var table = CreateTable(sheet);

                table.Sort(x => x.SortBy.ColumnNamed("Header1"));
                Assert.AreEqual(5, sheet.Cells[2, 1].Value);
                Assert.AreEqual(10, sheet.Cells[3, 1].Value);
                Assert.IsNotNull(table.SortState, "SortState was null");
                Assert.IsNotNull(table.SortState.SortConditions, "SortState.SortConditions was null");
                Assert.IsFalse(table.SortState.SortConditions.First().Descending, "First SortCondition was not descending");
            }
        }

        [TestMethod]
        public void TableSortByCustomList()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // header
                sheet.Cells[1, 1].Value = "Size";
                sheet.Cells[1, 2].Value = "Price";
                sheet.Cells[1, 3].Value = "Color";
                // row 1
                sheet.Cells[2, 1].Value = "M";
                sheet.Cells[2, 2].Value = 20;
                sheet.Cells[2, 3].Value = "Blue";
                // row 2
                sheet.Cells[3, 1].Value = "XL";
                sheet.Cells[3, 2].Value = 25;
                sheet.Cells[3, 3].Value = "Yellow";
                // row 3
                sheet.Cells[4, 1].Value = "S";
                sheet.Cells[4, 2].Value = 10;
                sheet.Cells[4, 3].Value = "Yellow";
                // row 4
                sheet.Cells[5, 1].Value = "L";
                sheet.Cells[5, 2].Value = 21;
                sheet.Cells[5, 3].Value = "Blue";
                // row 5
                sheet.Cells[6, 1].Value = "S";
                sheet.Cells[6, 2].Value = 20;
                sheet.Cells[6, 3].Value = "Blue";
                // row 6
                sheet.Cells[7, 1].Value = "S";
                sheet.Cells[7, 2].Value = 10;
                sheet.Cells[7, 3].Value = "Blue";

                var table = sheet.Tables.Add(sheet.Cells["A1:C7"], "myTable");

                table.Sort(x => x.SortBy.ColumnNamed("Size").UsingCustomList("S", "M", "L", "XL")
                                    .ThenSortBy.ColumnNamed("Price", eSortDirection.Descending)
                                    .ThenSortBy.Column(2).UsingCustomList("Blue", "Yellow"));
                
                
                Assert.AreEqual("S", sheet.Cells[2, 1].Value, $"First row, first col not 'S' but '{sheet.Cells[2, 1].Value}'");
                Assert.AreEqual(20, sheet.Cells[2, 2].Value, $"First row, second col not 20 but '{sheet.Cells[2, 2].Value}'");
                Assert.AreEqual("Blue", sheet.Cells[2, 3].Value, $"First row, third col not 'Blue' but '{sheet.Cells[2, 1].Value}'");

                Assert.AreEqual("S", sheet.Cells[3, 1].Value);
                Assert.AreEqual(10, sheet.Cells[3, 2].Value);
                Assert.AreEqual("Blue", sheet.Cells[3, 3].Value);

                Assert.AreEqual("S", sheet.Cells[4, 1].Value);
                Assert.AreEqual(10, sheet.Cells[4, 2].Value);
                Assert.AreEqual("Yellow", sheet.Cells[4, 3].Value);

                Assert.AreEqual("M", sheet.Cells[5, 1].Value);
                Assert.AreEqual("L", sheet.Cells[6, 1].Value);
                Assert.AreEqual("XL", sheet.Cells[7, 1].Value);
            }
        }
    }
}
