using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ExpandTests
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

        private void AddTestData()
        {
            var n = 1;
            for (var col = 1; col < 4; col++)
            {
                for (var row = 1; row < 5; row++)
                {
                    _sheet.Cells[row, col].Value = n++;
                }
            }
        }

        [TestMethod]
        public void ExpandShouldReturnValueErrorWhenRowTooSmall()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,1)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A10"].Value);
        }

        [TestMethod]
        public void ExpandShouldReturnValueErrorWhenColTooSmall()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,2,1)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A10"].Value);
        }

        [TestMethod]
        public void ExpandShouldReturnResultByMinimumArgs()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
        }

        [TestMethod]
        public void ExpandShouldReturnResultByOneCell()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1,2,2)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["A11"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["B10"].Value);
        }

        [TestMethod]
        public void ExpandShouldReturnResultByExpandedRow()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,3)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["A12"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["B12"].Value);
        }

        [TestMethod]
        public void ExpandShouldReturnResultByExpandedCol()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,,3)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["C10"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["C11"].Value);
        }

        [TestMethod]
        public void ExpandedCellsArePaddedWith4thArgument()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "EXPAND(A1:B2,,3,\"-\")";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual("-", _sheet.Cells["C10"].Value);
            Assert.AreEqual("-", _sheet.Cells["C11"].Value);
        }
    }
}
