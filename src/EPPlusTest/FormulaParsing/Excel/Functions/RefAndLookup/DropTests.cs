using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class DropTests
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
            for(var col = 1; col < 4; col++)
            {
                for(var row = 1; row < 5; row ++)
                {
                    _sheet.Cells[row, col].Value = n++;
                }
            }
        }

        [TestMethod]
        public void DropShouldReturnResultByRow()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4, 1)";
            _sheet.Calculate();
            Assert.AreEqual(2, _sheet.Cells["A10"].Value);
            Assert.AreEqual(3, _sheet.Cells["A11"].Value);
            Assert.AreEqual(4, _sheet.Cells["A12"].Value);
            Assert.AreEqual(6, _sheet.Cells["B10"].Value);
            Assert.AreEqual(7, _sheet.Cells["B11"].Value);
            Assert.AreEqual(8, _sheet.Cells["B12"].Value);
            Assert.AreEqual(10, _sheet.Cells["C10"].Value);
            Assert.AreEqual(11, _sheet.Cells["C11"].Value);
            Assert.AreEqual(12, _sheet.Cells["C12"].Value);
        }

        [TestMethod]
        public void DropShouldReturnResultByRowNegative()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4, -1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(7, _sheet.Cells["B12"].Value);
            Assert.AreEqual(9, _sheet.Cells["C10"].Value);
            Assert.AreEqual(10, _sheet.Cells["C11"].Value);
            Assert.AreEqual(11, _sheet.Cells["C12"].Value);
        }

        [TestMethod]
        public void DropShouldReturnResultByRowAndColNegative()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4,-1,-1)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(5, _sheet.Cells["B10"].Value);
            Assert.AreEqual(6, _sheet.Cells["B11"].Value);
            Assert.AreEqual(7, _sheet.Cells["B12"].Value);
        }

        [TestMethod]
        public void DropShouldReturnResultByRowAndCol()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4,1,1)";
            _sheet.Calculate();
            Assert.AreEqual(6, _sheet.Cells["A10"].Value);
            Assert.AreEqual(7, _sheet.Cells["A11"].Value);
            Assert.AreEqual(8, _sheet.Cells["A12"].Value);
            Assert.AreEqual(10, _sheet.Cells["B10"].Value);
            Assert.AreEqual(11, _sheet.Cells["B11"].Value);
            Assert.AreEqual(12, _sheet.Cells["B12"].Value);
        }

        [TestMethod]
        public void DropShouldReturnResultByNoArgs()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4,)";
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
        public void DropShouldReturnResultByColOnly()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4,,1)";
            _sheet.Calculate();
            Assert.AreEqual(5, _sheet.Cells["A10"].Value);
            Assert.AreEqual(6, _sheet.Cells["A11"].Value);
            Assert.AreEqual(7, _sheet.Cells["A12"].Value);
            Assert.AreEqual(8, _sheet.Cells["A13"].Value);
            Assert.AreEqual(9, _sheet.Cells["B10"].Value);
            Assert.AreEqual(10, _sheet.Cells["B11"].Value);
            Assert.AreEqual(11, _sheet.Cells["B12"].Value);
            Assert.AreEqual(12, _sheet.Cells["B13"].Value);
        }

        [TestMethod]
        public void DropShouldReturnResultByRow_InMemoryArray()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(SORT(A1:C4),1)";
            _sheet.Calculate();
            Assert.AreEqual(2, _sheet.Cells["A10"].Value);
            Assert.AreEqual(3, _sheet.Cells["A11"].Value);
            Assert.AreEqual(4, _sheet.Cells["A12"].Value);
            Assert.AreEqual(6, _sheet.Cells["B10"].Value);
            Assert.AreEqual(7, _sheet.Cells["B11"].Value);
            Assert.AreEqual(8, _sheet.Cells["B12"].Value);
            Assert.AreEqual(10, _sheet.Cells["C10"].Value);
            Assert.AreEqual(11, _sheet.Cells["C11"].Value);
            Assert.AreEqual(12, _sheet.Cells["C12"].Value);
        }

        [TestMethod]
        public void DropShouldHandleSingleArgument()
        {
            _sheet.Cells["A1"].Formula = "DROP(\"asdf\",1)";
            _sheet.Calculate();
            Assert.AreEqual("asdf", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void DropShouldReturnCalcWithSingleArgumentAndRowIs0()
        {
            _sheet.Cells["A1"].Formula = "DROP(\"asdf\",0)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.CalcError, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void DropShouldReturnCalcErrorIfRowIsTooLarge()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "DROP(A1:C4, 100)";
            _sheet.Calculate();
            Assert.AreEqual(ErrorValues.CalcError, _sheet.Cells["A10"].Value);
        }
    }
}
