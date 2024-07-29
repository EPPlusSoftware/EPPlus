using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

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

        [TestMethod]
        public void TrendTestUnevenSizes()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test where datapoints are equal but size is not");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 0;
                sheet.Cells["A5"].Value = 1;
                sheet.Cells["A6"].Value = 1;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["B3"].Value = 7;
                sheet.Cells["C2"].Value = 2;
                sheet.Cells["C3"].Value = 3;
                sheet.Cells["D2"].Value = 2;
                sheet.Cells["D3"].Value = 3;
                sheet.Cells["A8"].Formula = "TREND(A2:A6,B2:D3,,TRUE)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void TrendTestUnevenKnownXandNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test where input ranges knownX and Uneven X have different amount of columns");
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
                sheet.Cells["D2"].Value = 1;
                sheet.Cells["D3"].Value = 6;
                sheet.Cells["D4"].Value = 3;
                sheet.Cells["D5"].Value = 78;
                sheet.Cells["E2"].Value = 5;
                sheet.Cells["E3"].Value = 7;
                sheet.Cells["E4"].Value = 34;
                sheet.Cells["E5"].Value = 2;

                sheet.Cells["A8"].Formula = "TREND(A2:A5, B2:C5, D2:D5)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void TrendTestFewerNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with fewer new X observations");
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
                sheet.Cells["D2"].Value = 1;
                sheet.Cells["D3"].Value = 6;
                sheet.Cells["D4"].Value = 3;
                sheet.Cells["D5"].Value = 78;
                sheet.Cells["E2"].Value = 5;
                sheet.Cells["E3"].Value = 7;
                sheet.Cells["E4"].Value = 34;
                sheet.Cells["E5"].Value = 2;

                sheet.Cells["A8"].Formula = "TREND(A2:A5, B2:C5, D2:E3)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 8);
                Assert.AreEqual(4.217695212d, result1);
                Assert.AreEqual(62.20772199d, result2);

            }
        }
        [TestMethod]
        public void TrendTestMultipleRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple rows");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = -1;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["D1"].Value = 1;
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["B2"].Value = 1.23;
                sheet.Cells["C2"].Value = 2;
                sheet.Cells["D2"].Value = 6;
                sheet.Cells["A3"].Value = 423;
                sheet.Cells["B3"].Value = 33;
                sheet.Cells["C3"].Value = 6;
                sheet.Cells["D3"].Value = 3;
                sheet.Cells["A4"].Value = 7;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["C4"].Value = 7;
                sheet.Cells["D4"].Value = 78;
                sheet.Cells["A5"].Value = 1;
                sheet.Cells["B5"].Value = 4;
                sheet.Cells["C5"].Value = 7;
                sheet.Cells["D5"].Value = 3;

                sheet.Cells["A8"].Formula = "TREND(A1:D1, A2:D3, A4:D5, TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 9);
                Assert.AreEqual(1.820558632d, result1);
                Assert.AreEqual(1.744683966d, result2);
                Assert.AreEqual(1.80605155d, result3);
                Assert.AreEqual(3.033747918d, result4);

            }
        }

        [TestMethod]
        public void TrendTestCollinearRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with collinearity in original data-set");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["D1"].Value = 4;
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["B2"].Value = 1.23;
                sheet.Cells["C2"].Value = 2;
                sheet.Cells["D2"].Value = 6;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["C3"].Value = 7;
                sheet.Cells["D3"].Value = 8;
                sheet.Cells["A4"].Value = 7;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["C4"].Value = 7;
                sheet.Cells["D4"].Value = 78;
                sheet.Cells["A5"].Value = 1;
                sheet.Cells["B5"].Value = 4;
                sheet.Cells["C5"].Value = 7;
                sheet.Cells["D5"].Value = 3;

                sheet.Cells["A8"].Formula = "TREND(A1:D1, A2:D3, A4:D5, TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 9);
                Assert.AreEqual(-3d, result1);
                Assert.AreEqual(0d, result2);
                Assert.AreEqual(3d, result3);
                Assert.AreEqual(-1d, result4);


            }
        }


        [TestMethod]
        public void TrendTestDefaultX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with x-values omitted");
                sheet.Cells["A1"].Value = 232;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 21.121;
                sheet.Cells["D1"].Value = 332;
                sheet.Cells["A8"].Formula = "TREND(A1:D1)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 7);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 7);
                Assert.AreEqual(99.3121d, result1);
                Assert.AreEqual(131.1242d, result2);
                Assert.AreEqual(162.9363d, result3);
                Assert.AreEqual(194.7484d, result4);
            }
        }

        [TestMethod]
        public void TrendTestDefaultXCols()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with x-values omitted, column-based");
                sheet.Cells["A1"].Value = 232;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 21.121;
                sheet.Cells["A4"].Value = 332;
                sheet.Cells["A8"].Formula = "TREND(A1:A4)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 7);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 7);
                Assert.AreEqual(99.3121d, result1);
                Assert.AreEqual(131.1242d, result2);
                Assert.AreEqual(162.9363d, result3);
                Assert.AreEqual(194.7484d, result4);
            }
        }

        //[TestMethod]

        //public void WorkBookTest()
        //{
        //    var wbPath = "C:\\Users\\HannesAlm\\Downloads\\TrendTest.xlsx";
        //    using (ExcelPackage package = new ExcelPackage(new FileInfo(wbPath)))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        //        worksheet.Cells["O1"].Value = "EPPLUS RESULT";
        //        worksheet.Cells["O1"].Style.Font.Bold = true;
        //        worksheet.Cells["O2"].Formula = "TREND(A2:A1001, B2:F1001, G2:K1001, FALSE)";
        //        worksheet.Calculate();
        //        for (int i = 2; i < 1002; i++ )
        //        {
        //            var trendVal = (double)System.Math.Round((double)worksheet.Cells["N" + i].Value, 6);
        //            var epplusVal = (double)System.Math.Round((double)worksheet.Cells["O" + i].Value, 6);
        //            if (trendVal == epplusVal)
        //            {
        //                worksheet.Cells["P" + i].Value = "CORRECT";
        //            }
        //            else
        //            {
        //                worksheet.Cells["P" + i].Value = "INCORRECT";
        //            }
        //        }
        //        package.Save();
        //    }
        //}
    }
}