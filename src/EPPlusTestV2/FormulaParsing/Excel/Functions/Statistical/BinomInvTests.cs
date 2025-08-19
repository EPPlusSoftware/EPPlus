using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class BinomInvTests
    {

        [TestMethod]
        public void BinomInvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 6;
                sheet.Cells["A3"].Value = 0.5;
                sheet.Cells["A4"].Value = 0.75;

                sheet.Cells["B5"].Formula = "BINOM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(4d, result);
            }
        }

        [TestMethod]
        public void BinomInvShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 20;
                sheet.Cells["A3"].Value = 0.7;
                sheet.Cells["A4"].Value = 0.3;


                sheet.Cells["B5"].Formula = "BINOM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(13d, result);
            }
        }

        [TestMethod]
        public void BinomInvShouldReturnCorrectResultWhenTrailsIsZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Value = 0.7;
                sheet.Cells["A4"].Value = 0.3;


                sheet.Cells["B5"].Formula = "BINOM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(0d, result);
            }
        }

        [TestMethod]
        public void BinomInvShouldReturnCorrectResultWhenTrailsIsLarge()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 120;
                sheet.Cells["A3"].Value = 0.7;
                sheet.Cells["A4"].Value = 0.3;


                sheet.Cells["B5"].Formula = "BINOM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(81d, result);
            }
        }
    }
}