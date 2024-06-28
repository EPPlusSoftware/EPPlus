using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TrendTest : TestBase
    {
        [TestMethod]

        public void SimpleTrendTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Trend Test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 423;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = -1;
                sheet.Cells["B3"].Value = 1.23;
                sheet.Cells["B4"].Value = 33;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "TREND(A2:A5, B2:B5,,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                Assert.AreEqual(-20.27994279d, result1);
                Assert.AreEqual(8.606388047d, result2);
                Assert.AreEqual(420.1394512d, result3);
                Assert.AreEqual(31.53410356d, result4);
            }
        }
    
        [TestMethod]

        public void TrendWithNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Trend Test with newXs parameter");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 423;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = -1;
                sheet.Cells["B3"].Value = 1.23;
                sheet.Cells["B4"].Value = 33;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["C3"].Value = 2;
                sheet.Cells["C4"].Value = 6;
                sheet.Cells["C5"].Value = 7;
                sheet.Cells["A8"].Formula = "TREND(A2:A5, B2:B5,C2:C5,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 8);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 8);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                Assert.AreEqual(57.44112673d, result1);
                Assert.AreEqual(18.58059197d, result2);
                Assert.AreEqual(70.39463832d, result3);
                Assert.AreEqual(83.34814991d, result4);
            }
        }

        [TestMethod]

        public void TrendMultipleXsConstFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Trend Test with multiple X's");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 423;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = -1;
                sheet.Cells["B3"].Value = 1.23;
                sheet.Cells["B4"].Value = 33;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["C3"].Value = 2;
                sheet.Cells["C4"].Value = 6;
                sheet.Cells["C5"].Value = 7;
                sheet.Cells["A8"].Formula = "TREND(A2:A5, B2:C5,,FALSE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 8);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                Assert.AreEqual(-23.02271398d, result1);
                Assert.AreEqual(12.14808294d, result2);
                Assert.AreEqual(420.4802056d, result3);
                Assert.AreEqual(25.4194529d, result4);
            }
        }

        [TestMethod]

        public void TrendMultipleXsAndNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Trend Test with multiple X's");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 423;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = -1;
                sheet.Cells["B3"].Value = 1.23;
                sheet.Cells["B4"].Value = 33;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["C3"].Value = 2;
                sheet.Cells["C4"].Value = 6;
                sheet.Cells["C5"].Value = 7;
                sheet.Cells["D2"].Value = 2.73;
                sheet.Cells["D3"].Value = 0;
                sheet.Cells["D4"].Value = 498;
                sheet.Cells["D5"].Value = 284.453;
                sheet.Cells["E2"].Value = 453;
                sheet.Cells["E3"].Value = 1;
                sheet.Cells["E4"].Value = 34;
                sheet.Cells["E5"].Value = 3;
                sheet.Cells["A8"].Formula = "TREND(A2:A5,B2:C5,D2:E5,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 6);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 6);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 6);
                Assert.AreEqual(-1687.142955d, result1);
                Assert.AreEqual(6.393489381d, result2);
                Assert.AreEqual(6418.090568d, result3);
                Assert.AreEqual(3733.161974d, result4);
            }
        }
    }
}