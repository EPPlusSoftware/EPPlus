using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class NegBinomDotDistTests
    {

        [TestMethod]
        public void NegBinomDotDistShouldReturnCorrectResultWhenFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 10;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 0.25;
                sheet.Cells["A5"].Value = false;
                sheet.Cells["B5"].Formula = "NEGBINOM.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.0550487, result);
            }
        }

        [TestMethod]
        public void NegBinomDotDistShouldReturnCorrectResultWhenCumulative()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 10;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 0.25;
                sheet.Cells["A5"].Value = true;
                sheet.Cells["B5"].Formula = "NEGBINOM.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.3135141, result);
            }
        }

        [TestMethod]
        public void NegBinomDotDistShouldReturnCorrectResultWhenFalse2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 7;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 0.57;
                sheet.Cells["A5"].Value = false;
                sheet.Cells["B5"].Formula = "NEGBINOM.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.018122, result);
            }
        }

        [TestMethod]
        public void NegBinomDotDistShouldReturnCorrectResultWhenCumulative2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["A3"].Value = 7;
                sheet.Cells["A4"].Value = 0.87;
                sheet.Cells["A5"].Value = true;
                sheet.Cells["B5"].Formula = "NEGBINOM.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                Assert.AreEqual(0.9999948, result);
            }
        }

        [TestMethod]
        public void NegBinomDotDistShouldReturnCorrectResultWhenCumulativeArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["A3"].Value = 7;
                sheet.Cells["A4"].Value = 0.87;
                sheet.Cells["A5"].Value = true;
                sheet.Cells["B6"].Formula = "NEGBINOM.DIST(A1:A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B6"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B7"].Value, 7);

                Assert.AreEqual(0.7205567, result);
                Assert.AreEqual(0.9999948, result2);

            }
        }
    }
}
