using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class SortByTests
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;

        [TestInitialize]
        public void TestInitialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ValueErrorIfNoSortorderSupplied()
        {
            _sheet.Cells["A1"].Formula = "SORTBY(A2:C5,B2:B5,C2:C5)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ValueErrorIfByRangeHasMultipleRowsAndCols()
        {
            _sheet.Cells["A1"].Formula = "SORTBY(A2:C5,B2:C5)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ValueErrorIfTwoByRangesHasDifferentDirections()
        {
            _sheet.Cells["A1"].Formula = "SORTBY(A2:C5,C2:C5,1,D2:G2)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void SortByRowAscending_1SortCol()
        {
            _sheet.Cells["A1"].Value = "Bob";
            _sheet.Cells["B1"].Value = "Street 1";
            _sheet.Cells["A2"].Value = "Steve";
            _sheet.Cells["B2"].Value = "Street 2";
            _sheet.Cells["A3"].Value = "Phil";
            _sheet.Cells["B3"].Value = "Street 3";
            _sheet.Cells["C1"].Value = 25;
            _sheet.Cells["C2"].Value = 23;
            _sheet.Cells["C3"].Value = 21;
            _sheet.Cells["A4"].Formula = "SORTBY(A1:B3,C1:C3,1)";
            _sheet.Calculate();
            Assert.AreEqual("Phil", _sheet.Cells["A4"].Value);
            Assert.AreEqual("Steve", _sheet.Cells["A5"].Value);
            Assert.AreEqual("Bob", _sheet.Cells["A6"].Value);
        }

        [TestMethod]
        public void SortByRowAscending_2SortCols()
        {
            _sheet.Cells["A1"].Value = "Bob";
            _sheet.Cells["B1"].Value = "Street 1";
            _sheet.Cells["A2"].Value = "Steve";
            _sheet.Cells["B2"].Value = "Street 2";
            _sheet.Cells["A3"].Value = "Phil";
            _sheet.Cells["B3"].Value = "Street 3";
            _sheet.Cells["C1"].Value = 25;
            _sheet.Cells["C2"].Value = 25;
            _sheet.Cells["C3"].Value = 21;
            _sheet.Cells["D1"].Value = 1;
            _sheet.Cells["D2"].Value = 2;
            _sheet.Cells["D3"].Value = 3;
            _sheet.Cells["A4"].Formula = "SORTBY(A1:B3,C1:C3,1,D1:D3,1)";
            _sheet.Calculate();
            Assert.AreEqual("Phil", _sheet.Cells["A4"].Value);
            Assert.AreEqual("Street 3", _sheet.Cells["B4"].Value);
            Assert.AreEqual("Bob", _sheet.Cells["A5"].Value);
            Assert.AreEqual("Street 1", _sheet.Cells["B5"].Value);
            Assert.AreEqual("Steve", _sheet.Cells["A6"].Value);
            Assert.AreEqual("Street 2", _sheet.Cells["B6"].Value);
        }

        [TestMethod]
        public void SortByColAscending_1SortRow()
        {
            _sheet.Cells["A1"].Value = "Bob";
            _sheet.Cells["A2"].Value = "Street 1";
            _sheet.Cells["B1"].Value = "Steve";
            _sheet.Cells["B2"].Value = "Street 2";
            _sheet.Cells["C1"].Value = "Phil";
            _sheet.Cells["C2"].Value = "Street 3";
            _sheet.Cells["A3"].Value = 25;
            _sheet.Cells["B3"].Value = 23;
            _sheet.Cells["C3"].Value = 21;
            _sheet.Cells["A6"].Formula = "SORTBY(A1:C2,A3:C3,1)";
            _sheet.Calculate();
            Assert.AreEqual("Phil", _sheet.Cells["A6"].Value);
            Assert.AreEqual("Steve", _sheet.Cells["B6"].Value);
            Assert.AreEqual("Bob", _sheet.Cells["C6"].Value);
        }
    }
}
