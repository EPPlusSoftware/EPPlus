using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class GammaDotDistTests
    {
        [TestMethod]
        public void GammadotDistShouldReturnCorrectResultWhenCumulativeFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 10.00001131;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0.032639, result);
            }
        }

        [TestMethod]
        public void GammadotDistShouldReturnCorrectResultWhenCumulativeTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 10.00001131;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = true;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0.068094, result);
            }
        }

        [TestMethod]
        public void GammadotDistShouldReturnCorrectResultWhenCumulativeFalse2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 15.00005472;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0.068664, result);
            }
        }

        [TestMethod]
        public void GammadotDistShouldReturnZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 10.00001131;
                sheet.Cells["A3"].Value = 100;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0d, result);
            }
        }

        [TestMethod]
        public void GammadotDistShouldReturnCorrectResultBigXValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 50;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = false;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0.000026, result);
            }
        }


        [TestMethod]
        public void GammadotDistShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = true;

                sheet.Cells["B5"].Formula = "GAMMA.DIST(A2,A3,A4,A5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 6);
                Assert.AreEqual(0.593994, result);
            }
        }
    }
}
