using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class ChiSqDotDistTests
    {

        [TestMethod]
        public void ChiSqDotDistShouldReturnCorrectResultWhenCumulative()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 0.5;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = true;

                sheet.Cells["B4"].Formula = "CHISQ.DIST(A1,A2,A3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(0.52049988d, result);
            }
        }

        [TestMethod]
        public void ChiSqDotDistShouldReturnCorrectResultProbabilityDensity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = false;

                sheet.Cells["B4"].Formula = "CHISQ.DIST(A1,A2,A3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(0.20755375, result);
            }
        }

        [TestMethod]
        public void ChiSqDotDistShouldReturnCorrectResultProbabilityDensity2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 60;
                sheet.Cells["A2"].Value = 47;
                sheet.Cells["A3"].Value = false;

                sheet.Cells["B4"].Formula = "CHISQ.DIST(A1,A2,A3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(0.01500007, result);
            }
        }

        [TestMethod]
        public void ChiSqDotDistShouldReturnCorrectResultWhenCumulative2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 1000000;
                sheet.Cells["A2"].Value = 543235;
                sheet.Cells["A3"].Value = true;

                sheet.Cells["B4"].Formula = "CHISQ.DIST(A1,A2,A3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(1d, result);
            }
        }

        [TestMethod]
        public void ChiSqDotDistShouldReturnCorrectResultWhenCumulative3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 53456;
                sheet.Cells["A2"].Value = 53456;
                sheet.Cells["A3"].Value = true;

                sheet.Cells["B4"].Formula = "CHISQ.DIST(A1,A2,A3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(0.5008134, result);
            }
        }
    }
}