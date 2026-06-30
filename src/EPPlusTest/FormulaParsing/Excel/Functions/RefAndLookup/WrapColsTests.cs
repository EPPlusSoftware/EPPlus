using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class WrapColsTests
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

        private void AddRowVector()
        {
            // A1:F1 = 1..6
            for (var col = 1; col <= 6; col++)
            {
                _sheet.Cells[1, col].Value = col;
            }
        }

        private void AddColumnVector()
        {
            // A1:A6 = 1..6
            for (var row = 1; row <= 6; row++)
            {
                _sheet.Cells[row, 1].Value = row;
            }
        }

        [TestMethod]
        public void ShouldWrapRowVectorIntoColumnsOfThree()
        {
            AddRowVector();
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,3)";
            _sheet.Calculate();
            // Expected layout at A10:B12 (3 rows, 2 cols)
            //  1  4
            //  2  5
            //  3  6
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["B10"].Value);
            Assert.AreEqual(5, _sheet.Cells["B11"].Value);
            Assert.AreEqual(6, _sheet.Cells["B12"].Value);
        }

        [TestMethod]
        public void ShouldWrapColumnVectorIntoColumnsOfTwo()
        {
            AddColumnVector();
            _sheet.Cells["C10"].Formula = "WRAPCOLS(A1:A6,2)";
            _sheet.Calculate();
            // Expected layout at C10:E11 (2 rows, 3 cols)
            //  1  3  5
            //  2  4  6
            Assert.AreEqual(1, _sheet.Cells["C10"].Value);
            Assert.AreEqual(2, _sheet.Cells["C11"].Value);
            Assert.AreEqual(3, _sheet.Cells["D10"].Value);
            Assert.AreEqual(4, _sheet.Cells["D11"].Value);
            Assert.AreEqual(5, _sheet.Cells["E10"].Value);
            Assert.AreEqual(6, _sheet.Cells["E11"].Value);
        }

        [TestMethod]
        public void ShouldPadLastColumnWithNAByDefault()
        {
            AddRowVector();
            // 6 items, wrap_count = 4 -> 4 rows, 2 cols, last col has 2 padded cells
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,4)";
            _sheet.Calculate();
            // Expected
            //  1  5
            //  2  6
            //  3  #N/A
            //  4  #N/A
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            var b12 = _sheet.Cells["B12"].Value as ExcelErrorValue;
            var b13 = _sheet.Cells["B13"].Value as ExcelErrorValue;
            Assert.IsNotNull(b12);
            Assert.IsNotNull(b13);
            Assert.AreEqual(eErrorType.NA, b12.Type);
            Assert.AreEqual(eErrorType.NA, b13.Type);
        }

        [TestMethod]
        public void ShouldUseSuppliedPadValue()
        {
            AddRowVector();
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,4,0)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(0D, _sheet.Cells["B12"].Value);
            Assert.AreEqual(0D, _sheet.Cells["B13"].Value);
        }

        [TestMethod]
        public void ShouldReturnExactFitWithoutPadding()
        {
            AddRowVector();
            // 6 items, wrap_count = 3 -> exactly 2 cols, no padding
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,3,\"X\")";
            _sheet.Calculate();
            Assert.AreEqual(6, _sheet.Cells["B12"].Value);
            // Make sure the cell to the right of the spill is untouched
            Assert.IsNull(_sheet.Cells["C10"].Value);
        }

        [TestMethod]
        public void ShouldReturnValueErrorFor2dRange()
        {
            // 2x3 range is not a vector
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["C1"].Value = 3;
            _sheet.Cells["A2"].Value = 4;
            _sheet.Cells["B2"].Value = 5;
            _sheet.Cells["C2"].Value = 6;
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:C2,2)";
            _sheet.Calculate();
            var err = _sheet.Cells["A10"].Value as ExcelErrorValue;
            Assert.IsNotNull(err);
            Assert.AreEqual(eErrorType.Value, err.Type);
        }

        [TestMethod]
        public void ShouldReturnNumErrorWhenWrapCountIsZero()
        {
            AddRowVector();
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,0)";
            _sheet.Calculate();
            var err = _sheet.Cells["A10"].Value as ExcelErrorValue;
            Assert.IsNotNull(err);
            Assert.AreEqual(eErrorType.Num, err.Type);
        }

        [TestMethod]
        public void ShouldReturnNumErrorWhenWrapCountIsNegative()
        {
            AddRowVector();
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1:F1,-1)";
            _sheet.Calculate();
            var err = _sheet.Cells["A10"].Value as ExcelErrorValue;
            Assert.IsNotNull(err);
            Assert.AreEqual(eErrorType.Num, err.Type);
        }

        [TestMethod]
        public void ShouldWrapSingleCellAsOneByOne()
        {
            _sheet.Cells["A1"].Value = 42;
            _sheet.Cells["A10"].Formula = "WRAPCOLS(A1,1)";
            _sheet.Calculate();
            Assert.AreEqual(42, _sheet.Cells["A10"].Value);
        }
    }
}