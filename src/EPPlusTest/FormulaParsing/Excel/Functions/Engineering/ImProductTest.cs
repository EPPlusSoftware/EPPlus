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
    }
}
