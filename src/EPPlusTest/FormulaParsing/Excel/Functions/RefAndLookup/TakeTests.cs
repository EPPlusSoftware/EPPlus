using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class TakeTests
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
        public void TakeShouldReturnResultByRow()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4, 1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(9, _sheet.Cells["C10"].Value);
        }

        [TestMethod]
        public void TakeShouldReturnResultByRowNegative()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4, -1)";
            _sheet.Calculate();
            Assert.AreEqual(4, _sheet.Cells["A10"].Value);
            Assert.AreEqual(8, _sheet.Cells["B10"].Value);
            Assert.AreEqual(12, _sheet.Cells["C10"].Value);
        }

        [TestMethod]
        public void TakeShouldReturnResultByRowAndColNegative()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4,-1,-1)";
            _sheet.Calculate();
            Assert.AreEqual(12, _sheet.Cells["A10"].Value);
            Assert.IsNull(_sheet.Cells["A11"].Value, "A11 was not null");
            Assert.IsNull(_sheet.Cells["B10"].Value, "B10 was not null");
        }

        [TestMethod]
        public void TakeShouldReturnResultByRowAndCol()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4,1,1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.IsNull(_sheet.Cells["A11"].Value, "A11 was not null");
            Assert.IsNull(_sheet.Cells["B10"].Value, "B10 was not null");
        }

        [TestMethod]
        public void TakeShouldReturnResultByNoArgs()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4,)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(7, _sheet.Cells["B12"].Value);
            Assert.AreEqual(8, _sheet.Cells["B13"].Value);
            Assert.AreEqual(10, _sheet.Cells["C11"].Value);
            Assert.AreEqual(11, _sheet.Cells["C12"].Value);
            Assert.AreEqual(12, _sheet.Cells["C13"].Value);
        }

        [TestMethod]
        public void TakeShouldReturnResultByColOnly()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4,,1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
        }

        [TestMethod]
        public void TakeShouldReturnResultByRow_InMemoryArray()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(SORT(A1:C4),1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(9, _sheet.Cells["C10"].Value);
        }

        [TestMethod]
        public void TakeShouldReturnResultWhenRowsAndColsAreLargerThanRangeSize()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TAKE(A1:C4,100,100)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(7, _sheet.Cells["B12"].Value);
            Assert.AreEqual(8, _sheet.Cells["B13"].Value);
            Assert.AreEqual(10, _sheet.Cells["C11"].Value);
            Assert.AreEqual(11, _sheet.Cells["C12"].Value);
            Assert.AreEqual(12, _sheet.Cells["C13"].Value);
        }
    }
}
