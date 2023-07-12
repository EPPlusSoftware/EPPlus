using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class PoissonTest : TestBase
    {

        [TestMethod]
        public void PoissonPDFTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Poisson should return correct result: ");
                sheet.Cells["A1"].Formula = "POISSON(6,9,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.091090319d, result);
            }
        }

        [TestMethod] 
        public void PoissonCdfTest() 
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Poisson should return correct result: ");
                sheet.Cells["A1"].Formula = "POISSON(11,8,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.888075999, result);
            }
        }

        [TestMethod]
        public void PoissonNegativeX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Should return an error: ");
                sheet.Cells["A1"].Formula = "POISSON(-11,8,TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(result, ExcelErrorValue.Create(eErrorType.Num));
            }
        }

        [TestMethod]
        public void PoissonNegativeMean()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Should return an error: ");
                sheet.Cells["A1"].Formula = "POISSON(11,-8,TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(result, ExcelErrorValue.Create(eErrorType.Num));
            }
        }

        [TestMethod]
        public void PoissonTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Poisson should return correct result: ");
                sheet.Cells["A1"].Formula = "POISSON(19.2345,8.2453,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.999631483d, result);
            }
        }

        [TestMethod]
        public void PossionTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Poisson should return correct result: ");
                sheet.Cells["A1"].Formula = "POISSON(2.67,8.24,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.008958554d, result);
            }
        }

        [TestMethod]
        public void PoissonTruncateX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Same as PoissonTest2 but x is lowered to nearest integer: ");
                sheet.Cells["A1"].Formula = "POISSON(2,8.24,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.008958554d, result);
            }
        }
    }
}
