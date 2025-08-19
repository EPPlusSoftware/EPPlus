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
    public class ChisqtestTests
    {

        [TestMethod]
        public void ChisqTestShouldReturnCorrectResult ()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["B4"].Formula = "CHISQ.TEST(A1:A3,B1:B3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 9);
                Assert.AreEqual(0.062349477d, result);
            }
        }

        [TestMethod]
        public void ChisqTestDifferentRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with different ranges");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B4"].Formula = "CHISQ.TEST(A1:A3, B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);
            } 
        }

        [TestMethod]
        public void ChisqTestOneRowSeveralColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with one row and more than one columns");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["C2"].Value = 6;
                sheet.Cells["A3"].Formula = "CHISQ.TEST(A1:C1, A2:C2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A3"].Value, 9);
                Assert.AreEqual(0.062349477d, result);
            }
        }

        [TestMethod]
        public void ChisqTestOneColumnSeveralRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with one column but several rows");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["B4"].Formula = "CHISQ.TEST(A1:A3,B1:B3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 9);
                Assert.AreEqual(0.062349477d, result);
            }
        }

        [TestMethod]
        public void ChisqTestSeveralRowsSeveralColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with two matrices");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["B2"].Value = 4;
                sheet.Cells["D1"].Value = 5;
                sheet.Cells["D2"].Value = 6;
                sheet.Cells["E1"].Value = 7;
                sheet.Cells["E2"].Value = 8;
                sheet.Cells["B4"].Formula = "CHISQ.TEST(A1:B2,D1:E2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B4"].Value, 8);
                Assert.AreEqual(0.00144115d, result);
            }
        }

        [TestMethod]
        public void ChisqTestOneDatapoint()
        {
            using (var package = new ExcelPackage()) 
            {
                var sheet = package.Workbook.Worksheets.Add("Test with one datapoint");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B4"].Formula = "CHISQ.TEST(A1, A2)";
                sheet.Calculate();
                var result = sheet.Cells["B4"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);


            }
        }

        [TestMethod]
        public void ChisqTestMicrosoftExample()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test data");
                sheet.Cells["A2"].Value = 58;
                sheet.Cells["A3"].Value = 11;
                sheet.Cells["A4"].Value = 10;
                sheet.Cells["B2"].Value = 35;
                sheet.Cells["B3"].Value = 25;
                sheet.Cells["B4"].Value = 23;
                sheet.Cells["A6"].Value = 45.35;
                sheet.Cells["A7"].Value = 17.56;
                sheet.Cells["A8"].Value = 16.09;
                sheet.Cells["B6"].Value = 47.65;
                sheet.Cells["B7"].Value = 18.44;
                sheet.Cells["B8"].Value = 16.91;
                sheet.Cells["B15"].Formula = "CHISQ.TEST(A2:B4, A6:B8)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B15"].Value, 7);
                Assert.AreEqual(0.0003082d, result);
            }
        }

    }
}
