using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ImProductTest
    {
        [TestMethod]
        public void ImProductShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"3+5i\", \"2+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-14+22i", result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnCorrectResult_3Numbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"3+5i\", \"2+4i\", \"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-130+10i", result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnCorrectResult_onlyReal()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"3\", \"2+4i\", \"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-30+60i", result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnCorrectResult_onlyImag()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"i\", \"2+4i\", \"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-20-10i", result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnCorrectResult_WithRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "3+5i";
                sheet.Cells["A2"].Value = "2+4i";
                sheet.Cells["A3"].Value = "3+4i";
                sheet.Cells["A4"].Formula = "IMPRODUCT(A1:A3)";
                sheet.Calculate();
                var result = sheet.Cells["A4"].Value;
                Assert.AreEqual("-130+10i", result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnError_WithDifferentSuffixes()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"i\", \"2+4j\", \"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void ImProductShouldReturnError_WhenInvalidInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPRODUCT(\"i\", \"2p+4j\", \"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
    }
}
