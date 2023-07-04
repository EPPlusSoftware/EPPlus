using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{

    [TestClass]
    public class TDistTest : TestBase
    {
        [TestMethod]

        public void TDistWithFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST(5,2,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.007127781, result);
            }
        }

        [TestMethod]
        public void TDistWithTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST(5,2,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.981125224, result);
            }
        }

        [TestMethod]
        public void TDistWithNegativeXAndTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel, with negative numeric value x");
                sheet.Cells["A1"].Formula = "T.DIST(-8,2,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.007634036, result);
            }
        }

        [TestMethod]
        public void TDistWithNegativeXAndFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel, with negative numeric value x");
                sheet.Cells["A1"].Formula = "T.DIST(-12.879,3.6978,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.000116002, result);
            }
        }

        [TestMethod]
        public void TDistDegreesOfFreedomLessThanOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #DIV/0! if df is less than 1: ");
                sheet.Cells["A1"].Formula = "T.DIST(-12.879,0.5,FALSE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

        [TestMethod]
        public void TDistIncorrectX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #VALUE! if X or df is not numeric: ");
                sheet.Cells["A1"].Formula = "T.DIST(\"hgjk\",2,TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void TDistCrazyNegativeInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result");
                sheet.Cells["A1"].Formula = "T.DIST(-17.879,9.321,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 13);
                Assert.AreEqual(1.21734E-08, result);
            }
        }

        [TestMethod]
        public void TDistCrazyPositiveInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result");
                sheet.Cells["A1"].Formula = "T.DIST(19.4856,7.53432,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.999999883, result);
            }
        }

        [TestMethod]
        public void TDistCrazyPDF()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result");
                sheet.Cells["A1"].Formula = "T.DIST(21.235756549,5.9999642353,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 12);
                Assert.AreEqual(5.00577E-07, result);
            }
        }

        [TestMethod]
        public void TDistWithZeroEqualsFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result");
                sheet.Cells["A1"].Formula = "T.DIST(21.235756549,5.9999642353,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 12);
                Assert.AreEqual(5.00577E-07, result);
            }
        }

        [TestMethod]
        public void TDistWithOneEqualsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result");
                sheet.Cells["A1"].Formula = "T.DIST(5,2,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.981125224, result);
            }
        }
    }
}
