using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class WeibullDistTest : TestBase
    {

        [TestMethod]
        public void WeibullDistCDF() 
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with cumulative distribution function: ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(105,20,100, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.92958139d, result);
            }
        }

        [TestMethod]
        public void WeibullDistPDF()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with cumulative distribution function: ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(105,20,100, FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.035588864d, result);
            }
        }

        [TestMethod]
        public void WeibullDistPDFAlphaIsOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with cumulative distribution function: ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(105,1,100, FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.003499377d, result);
            }
        }

        [TestMethod]
        public void WeibullDistCDFAlphaIsOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with cumulative distribution function: ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(105,1,100, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.650062251d, result);
            }
        }

        [TestMethod]
        public void WeibullDistTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(9.42,7.342,9, FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.269250417d, result);
            }
        }

        [TestMethod]
        public void WeibullDistTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return correct result ");
                sheet.Cells["A1"].Formula = "WEIBULL.DIST(9.42,7.342,9, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.75285053d, result);
            }
        }

        [TestMethod]
        public void WeibullTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Calling WEIBULL instead of WEIBULL.DIST ");
                sheet.Cells["A1"].Formula = "WEIBULL(9.42,7.342,9, TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.75285053d, result);
            }
        }
    }   
}
