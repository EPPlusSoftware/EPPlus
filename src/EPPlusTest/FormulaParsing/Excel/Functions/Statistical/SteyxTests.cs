using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class SteyxTests : TestBase
    {
        [TestMethod]
        public void SteyxShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["C1"].Value = 1;
                sheet.Cells["C2"].Value = 2;
                sheet.Cells["C3"].Value = 3;
                sheet.Cells["C4"].Value = 4;
                sheet.Cells["D1"].Value = 5;
                sheet.Cells["D2"].Value = 6;
                sheet.Cells["D3"].Value = 89;
                sheet.Cells["D4"].Value = 8;
                sheet.Cells["D5"].Formula = "STEYX(C1:C4,D1:D4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["D5"].Value, 9);
                Assert.AreEqual(1.514517145d, result);
            }
        }

        [TestMethod]
        public void SteyxTestDifferentRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test where knownY and knownX have different amount of data points.");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 9;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["B1"].Value = 7;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 5;
                sheet.Cells["B5"].Formula = "STEYX(A1:A5,B1:B4)";
                sheet.Calculate();
                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);
            }
        }

        [TestMethod]
        public void SteyxTestLessThanThreeDatapoints()
        {
            using ( var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test where theres less than three datapoints. Should return #DIV/0!");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = 7;
                sheet.Cells["B1"].Value = 9;
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["B5"].Formula = "STEYX(A1:A2,B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

        [TestMethod]
        public void SteyxTestContainingTextAndEmptyCells()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with a bunch of non-numeric cells");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = "Bicycle";
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 9;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["A6"].Value = 7;
                sheet.Cells["A7"].Value = 5;
                sheet.Cells["A8"].Value = "";
                sheet.Cells["B1"].Value = "";
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = "Spaceship";
                sheet.Cells["B5"].Value = 9;
                sheet.Cells["B6"].Value = 8;
                sheet.Cells["B7"].Value = 6;
                sheet.Cells["B8"].Value = 7;
                sheet.Cells["B9"].Formula = "STEYX(A1:A8,B1:B8)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                Assert.AreEqual(1.154700538d, result);
            }
        }

        [TestMethod]
        public void SteyxTestContainingZeros()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with ranges containing zeros");
                sheet.Cells["A1"].Value = 15;
                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 0;
                sheet.Cells["B1"].Value = 0;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 50;
                sheet.Cells["B4"].Value = 1;
                sheet.Cells["B5"].Formula = "STEYX(A1:A4,B1:B4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B5"].Value, 9);
                Assert.AreEqual(8.810144175d, result);
            }
        }

        [TestMethod]
        public void SteyxTestContainingTextEmptyCellsAndZeros()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with a bunch of non-numeric cells");
                sheet.Cells["A1"].Value = 4;
                sheet.Cells["A2"].Value = "Bicycle";
                sheet.Cells["A3"].Value = 0;
                sheet.Cells["A4"].Value = 9;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["A6"].Value = 7;
                sheet.Cells["A7"].Value = 0;
                sheet.Cells["A8"].Value = "";
                sheet.Cells["B1"].Value = "";
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = "Spaceship";
                sheet.Cells["B5"].Value = 0;
                sheet.Cells["B6"].Value = 8;
                sheet.Cells["B7"].Value = 6;
                sheet.Cells["B8"].Value = 7;
                sheet.Cells["B9"].Formula = "STEYX(A1:A8,B1:B8)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                Assert.AreEqual(5.249554565d, result);
            }
        }
    }
}
