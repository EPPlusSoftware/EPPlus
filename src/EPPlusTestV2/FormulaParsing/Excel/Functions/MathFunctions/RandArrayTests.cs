using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class RandArrayTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup() 
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldCreateOneNumberByDefault()
        {
            _sheet.Cells["A1"].Formula = "RANDARRAY()";
            _sheet.Calculate();
            Assert.IsInstanceOfType(_sheet.Cells["A1"].Value, typeof(double));
            var d = (double)_sheet.Cells["A1"].Value;
            Assert.IsTrue(d > 0);
            Assert.IsTrue(d < 1);
        }

        [TestMethod]
        public void ShouldCreateTwoRows()
        {
            _sheet.Cells["A1"].Formula = "RANDARRAY(2)";
            _sheet.Calculate();
            Assert.IsInstanceOfType(_sheet.Cells["A1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["A2"].Value, typeof(double));
            var a1 = (double)_sheet.Cells["A1"].Value;
            Assert.IsTrue(a1 > 0);
            Assert.IsTrue(a1 < 1);
            var a2 = (double)_sheet.Cells["A2"].Value;
            Assert.IsTrue(a2 > 0);
            Assert.IsTrue(a2 < 1);
        }

        [TestMethod]
        public void ShouldCreateTwoRowsAndTwoCols()
        {
            _sheet.Cells["A1"].Formula = "RANDARRAY(2,2)";
            _sheet.Calculate();
            Assert.IsInstanceOfType(_sheet.Cells["A1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["A2"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B2"].Value, typeof(double));
            var a1 = (double)_sheet.Cells["A1"].Value;
            Assert.IsTrue(a1 > 0);
            Assert.IsTrue(a1 < 1);
            var a2 = (double)_sheet.Cells["A2"].Value;
            Assert.IsTrue(a2 > 0);
            Assert.IsTrue(a2 < 1);
            var b1 = (double)_sheet.Cells["B1"].Value;
            Assert.IsTrue(b1 > 0);
            Assert.IsTrue(b1 < 1);
        }

        [TestMethod]
        public void ShouldCreateTwoRowsAndTwoCols_Between5and10()
        {
            _sheet.Cells["A1"].Formula = "RANDARRAY(2,2,5,10)";
            _sheet.Calculate();
            Assert.IsInstanceOfType(_sheet.Cells["A1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["A2"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B2"].Value, typeof(double));
            var a1 = (double)_sheet.Cells["A1"].Value;
            Assert.IsTrue(a1 > 5);
            Assert.IsTrue(a1 < 10);
            var a2 = (double)_sheet.Cells["A2"].Value;
            Assert.IsTrue(a2 > 5);
            Assert.IsTrue(a2 < 10);
            var b1 = (double)_sheet.Cells["B1"].Value;
            Assert.IsTrue(b1 > 5);
            Assert.IsTrue(b1 < 10);
        }

        [TestMethod]
        public void ShouldCreateTwoRowsAndTwoCols_Between5and10_Int()
        {
            _sheet.Cells["A1"].Formula = "RANDARRAY(2,2,5,10,TRUE)";
            _sheet.Calculate();
            Assert.IsInstanceOfType(_sheet.Cells["A1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["A2"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B1"].Value, typeof(double));
            Assert.IsInstanceOfType(_sheet.Cells["B2"].Value, typeof(double));
            var a1 = (double)_sheet.Cells["A1"].Value;
            Assert.IsTrue(a1 >= 5);
            Assert.IsTrue(a1 <= 10);
            Assert.IsTrue(a1 - System.Math.Floor(a1) < System.Math.Pow(10, -10));
            var a2 = (double)_sheet.Cells["A2"].Value;
            Assert.IsTrue(a2 >= 5);
            Assert.IsTrue(a2 <= 10);
            Assert.IsTrue(a2 - System.Math.Floor(a2) < System.Math.Pow(10, -10));
            var b1 = (double)_sheet.Cells["B1"].Value;
            Assert.IsTrue(b1 >= 5);
            Assert.IsTrue(b1 <= 10);
        }
    }
}
