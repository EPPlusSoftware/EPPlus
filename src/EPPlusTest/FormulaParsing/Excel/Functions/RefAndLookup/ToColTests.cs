using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class ToColTests
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
        public void ShouldReturnCorrectResultWithRange()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(5, _sheet.Cells["A11"].Value);
            Assert.AreEqual(9, _sheet.Cells["A12"].Value);
            Assert.AreEqual(2, _sheet.Cells["A13"].Value);
            Assert.AreEqual(6, _sheet.Cells["A14"].Value);
            Assert.AreEqual(10, _sheet.Cells["A15"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithIgnore0()
        {
            AddTestData();
            _sheet.Cells["A1"].Value = null;
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2,0)";
            _sheet.Calculate();
            Assert.AreEqual(0D, _sheet.Cells["A10"].Value);
            Assert.AreEqual(5, _sheet.Cells["A11"].Value);
            Assert.AreEqual(9, _sheet.Cells["A12"].Value);
            Assert.AreEqual(2, _sheet.Cells["A13"].Value);
            Assert.AreEqual(6, _sheet.Cells["A14"].Value);
            Assert.AreEqual(10, _sheet.Cells["A15"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithIgnore1()
        {
            AddTestData();
            _sheet.Cells["A1"].Value = null;
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2,1)";
            _sheet.Calculate();
            Assert.AreEqual(5, _sheet.Cells["A10"].Value);
            Assert.AreEqual(9, _sheet.Cells["A11"].Value);
            Assert.AreEqual(2, _sheet.Cells["A12"].Value);
            Assert.AreEqual(6, _sheet.Cells["A13"].Value);
            Assert.AreEqual(10, _sheet.Cells["A14"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithIgnore2()
        {
            AddTestData();
            _sheet.Cells["A1"].Formula = "1/0";
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2,2)";
            _sheet.Calculate();
            Assert.AreEqual(5, _sheet.Cells["A10"].Value);
            Assert.AreEqual(9, _sheet.Cells["A11"].Value);
            Assert.AreEqual(2, _sheet.Cells["A12"].Value);
            Assert.AreEqual(6, _sheet.Cells["A13"].Value);
            Assert.AreEqual(10, _sheet.Cells["A14"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithIgnore3()
        {
            AddTestData();
            _sheet.Cells["A1"].Formula = "1/0";
            _sheet.Cells["B1"].Value = null;
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2,3)";
            _sheet.Calculate();
            Assert.AreEqual(9, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(6, _sheet.Cells["A12"].Value);
            Assert.AreEqual(10, _sheet.Cells["A13"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithScanByColumnTrue()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C4,,TRUE)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(2, _sheet.Cells["A11"].Value);
            Assert.AreEqual(3, _sheet.Cells["A12"].Value);
            Assert.AreEqual(4, _sheet.Cells["A13"].Value);
            Assert.AreEqual(5, _sheet.Cells["A14"].Value);
            Assert.AreEqual(6, _sheet.Cells["A15"].Value);
        }

        [TestMethod]
        public void ShouldReturnCorrectResultWithScanByColumnFalse()
        {
            AddTestData();
            _sheet.Cells["A10"].Formula = "TOCOL(A1:C2,,FALSE)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A10"].Value);
            Assert.AreEqual(5, _sheet.Cells["A11"].Value);
            Assert.AreEqual(9, _sheet.Cells["A12"].Value);
            Assert.AreEqual(2, _sheet.Cells["A13"].Value);
            Assert.AreEqual(6, _sheet.Cells["A14"].Value);
            Assert.AreEqual(10, _sheet.Cells["A15"].Value);
        }
    }
}
