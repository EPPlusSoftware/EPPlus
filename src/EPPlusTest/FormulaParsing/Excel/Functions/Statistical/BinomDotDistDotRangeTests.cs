using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class BinomDotDistDotRangeTests
    {

        [TestMethod]
        public void BinomDistRangeShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 60;
                sheet.Cells["A3"].Value = 0.75;
                sheet.Cells["A4"].Value = 48;


                sheet.Cells["B5"].Formula = "BINOM.DIST.RANGE(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 3);
                Assert.AreEqual(0.084, result);
            }
        }

        [TestMethod]
        public void BinomDistRangeShouldReturnCorrectResultWithAllInputs()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 60;
                sheet.Cells["A3"].Value = 0.75;
                sheet.Cells["A4"].Value = 45;
                sheet.Cells["A5"].Value = 50;


                sheet.Cells["B5"].Formula = "BINOM.DIST.RANGE(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 3);
                Assert.AreEqual(0.524, result);
            }
        }

        [TestMethod]
        public void BinomDistRangeShouldReturnCorrectResultWithAllInputs2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 155;
                sheet.Cells["A3"].Value = 0.47;
                sheet.Cells["A4"].Value = 100;
                sheet.Cells["A5"].Value = 150;


                sheet.Cells["B5"].Formula = "BINOM.DIST.RANGE(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(8.5E-06, result);
            }
        }

        [TestMethod]
        public void BinomDistRangeShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 8;
                sheet.Cells["A3"].Value = 0.85;
                sheet.Cells["A4"].Value = 4;


                sheet.Cells["B5"].Formula = "BINOM.DIST.RANGE(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.0184986, result);
            }
        }
    }
}

