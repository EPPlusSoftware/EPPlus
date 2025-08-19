using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class BinomDistTests
    {

        [TestMethod]
        public void BinomDistShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 6;
                sheet.Cells["A3"].Value = 10;
                sheet.Cells["A4"].Value = 0.5;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.2050781, result);
            }
        }


        [TestMethod]
        public void BinomDistShouldReturnCorrectResultWhenCumulative()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 6;
                sheet.Cells["A3"].Value = 10;
                sheet.Cells["A4"].Value = 0.5;
                sheet.Cells["A5"].Value = true;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.828125, result);
            }
        }


        [TestMethod]
        public void BinomDistShouldReturnCorrectResultWhenCumulative2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 5;
                sheet.Cells["A3"].Value = 14;
                sheet.Cells["A4"].Value = 0.7;
                sheet.Cells["A5"].Value = true;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.0082885, result);
            }
        }

        [TestMethod]
        public void BinomDistShouldReturnCorrectResultWhenNotCumulative()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 8;
                sheet.Cells["A4"].Value = 0.3;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.1361367, result);
            }
        }

        [TestMethod]
        public void BinomDistShouldReturnCorrectResultdecimalInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4.432;
                sheet.Cells["A3"].Value = 8.243534;
                sheet.Cells["A4"].Value = 0.3;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.1361367, result);
            }
        }

        [TestMethod]
        public void BinomDistShouldReturnCorrectResultdecimalInput2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4.832;
                sheet.Cells["A3"].Value = 8.243534;
                sheet.Cells["A4"].Value = 0.3;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "BINOMDIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.1361367, result);
            }
        }
    }
}