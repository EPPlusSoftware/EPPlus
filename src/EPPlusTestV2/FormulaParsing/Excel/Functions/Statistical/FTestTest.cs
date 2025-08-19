using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class FTestTest : TestBase
    {

        [TestMethod]
        public void FTestShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("F-test: ");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 9;
                sheet.Cells["A6"].Value = 1;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Value = 4;
                sheet.Cells["A9"].Value = 5;
                sheet.Cells["B1"].Value = 6;
                sheet.Cells["B2"].Value = 19;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 14;
                sheet.Cells["B6"].Value = 4;
                sheet.Cells["B7"].Value = 5;
                sheet.Cells["B8"].Value = 17;
                sheet.Cells["B9"].Value = 1;
                sheet.Cells["B10"].Formula = "F.TEST(A1:A9,B1:B9)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                Assert.AreEqual(0.012763405d, result);
            }
        }

        [TestMethod]
        public void FTestWithUnevenRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("F-test: ");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = "gfd";
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 9;
                sheet.Cells["A6"].Value = 1;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Value = "";
                sheet.Cells["A9"].Value = 5;
                sheet.Cells["B1"].Value = 6;
                sheet.Cells["B2"].Value = 19;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 14;
                sheet.Cells["B6"].Value = 4;
                sheet.Cells["B7"].Value = 5;
                sheet.Cells["B8"].Value = 17;
                sheet.Cells["B9"].Value = 1;
                sheet.Cells["B10"].Formula = "F.TEST(A1:A9,B1:B9)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                Assert.AreEqual(0.057949721, result);
            }
        }

        [TestMethod]
        public void FTestMatrix()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("F-test: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 87;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 7;
                sheet.Cells["B1"].Value = 7;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["B3"].Value = 876;
                sheet.Cells["B4"].Value = 90;
                sheet.Cells["C1"].Value = 8;
                sheet.Cells["C2"].Value = 65;
                sheet.Cells["C3"].Value = 86;
                sheet.Cells["C4"].Value = 345;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 56;
                sheet.Cells["A7"].Value = 7;
                sheet.Cells["A8"].Value = 98;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["B6"].Value = 90;
                sheet.Cells["B7"].Value = 5;
                sheet.Cells["B8"].Value = 45;
                sheet.Cells["C5"].Value = 36;
                sheet.Cells["C6"].Value = 786;
                sheet.Cells["C7"].Value = 3;
                sheet.Cells["C8"].Value = 86;
                sheet.Cells["B15"].Formula = "F.TEST(A1:C4,A5:C8)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B15"].Value, 9);
                Assert.AreEqual(0.636614303d, result);
            }
        }

        [TestMethod]
        public void FTestArrayLessThanTwo()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect array: ");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = "";
                sheet.Cells["A3"].Value = "";
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 93;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["B5"].Formula = "F.TEST(A1:A3,B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

        [TestMethod]
        public void FTestVarianceIsZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect array: ");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["B5"].Formula = "F.TEST(A1:A3,B1:B3)";
                sheet.Calculate();
                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }
    }
}
