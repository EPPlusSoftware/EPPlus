using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ChooseColsTests
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
        public void ShouldReturnArrayWithSelectedCols()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "CHOOSECOLS(A1:C4,1,3, 1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(9, _sheet.Cells["B10"].Value);
            Assert.AreEqual(10, _sheet.Cells["B11"].Value);
            Assert.AreEqual(11, _sheet.Cells["B12"].Value);
            Assert.AreEqual(12, _sheet.Cells["B13"].Value);
            Assert.AreEqual(1, _sheet.Cells["C10"].Value);
            Assert.AreEqual(2, _sheet.Cells["C11"].Value);
            Assert.AreEqual(3, _sheet.Cells["C12"].Value);
            Assert.AreEqual(4, _sheet.Cells["C13"].Value);
        }

        [TestMethod]
        public void ShouldReturnArrayWithLastColIfNegativeColNumber()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "CHOOSECOLS(A1:C4,-1)";
            _sheet.Calculate();
            Assert.AreEqual(9, _sheet.Cells["A10"].Value);
            Assert.AreEqual(10, _sheet.Cells["A11"].Value);
            Assert.AreEqual(11, _sheet.Cells["A12"].Value);
            Assert.AreEqual(12, _sheet.Cells["A13"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithSingleCell()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "CHOOSECOLS(A1,1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithSingleCellAndMultiCols()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "CHOOSECOLS(A1,1,1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(1, _sheet.Cells["B10"].Value);
        }

        [TestMethod]
        public void ShouldReturnValueErrorWithSingleCellAndTooHighIndex()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "CHOOSECOLS(A1,2)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.ValueError, _sheet.Cells["A10"].Value);
        }
    }
}
