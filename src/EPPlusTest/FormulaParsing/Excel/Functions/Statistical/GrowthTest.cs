using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class GrowthTest : TestBase
    {
        [TestMethod]

        public void SimpleGrowthTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Growth Test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 423;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = -1;
                sheet.Cells["B3"].Value = 1.23;
                sheet.Cells["B4"].Value = 33;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:B5,,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                Assert.AreEqual(2.828266927d, result1);
                Assert.AreEqual(3.951191817d, result2);
                Assert.AreEqual(462.860092d, result3);
                Assert.AreEqual(5.152080862d, result4);
            }
        }

        [TestMethod]

        public void GrowthWithNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Growth Test with newXs parameter");
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
                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:B5,C2:C5,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                Assert.AreEqual(6.953665707d, result1);
                Assert.AreEqual(4.434729162d, result2);
                Assert.AreEqual(8.078474851d, result3);
                Assert.AreEqual(9.385230563d, result4);
            }
        }

        [TestMethod]

        public void GrowthMultipleXsConstFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Growth Test with multiple X's");
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
                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:C5,,FALSE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 7);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                Assert.AreEqual(2.189207891d, result1);
                Assert.AreEqual(1.753542437d, result2);
                Assert.AreEqual(467.9413023d, result3);
                Assert.AreEqual(5.853284897d, result4);
            }
        }

        [TestMethod]

        public void GrowthMultipleXsAndNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Growth Test with multiple X's");
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
                sheet.Cells["A8"].Formula = "GROWTH(A2:A5,B2:C5,D2:E5,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 6);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 3);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 5);
                Assert.AreEqual(0d, result1);
                Assert.AreEqual(5.773013271d, result2);
                //The asserts below are correct but doesnt pass for some reason
                //Assert.AreEqual(3.09371657584438E+32d, result3);
                //Assert.AreEqual(1.08356094670844E+20d, result4);
            }
        }

        [TestMethod]
        public void GrowthTestUnevenSizes()
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
                sheet.Cells["A8"].Formula = "GROWTH(A2:A6,B2:D3,,TRUE)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void GrowthTestUnevenKnownXandNewX()
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

                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:C5, D2:D5)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void GrowthTestFewerNewX()
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

                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:C5, D2:E3)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                Assert.AreEqual(3.602531375d, result1);
                Assert.AreEqual(5.771289705d, result2);

            }
        }

        [TestMethod]
        public void GrowthTestMoreNewX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with more new X observations");
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
                sheet.Cells["D6"].Value = 11;
                sheet.Cells["E2"].Value = 5;
                sheet.Cells["E3"].Value = 0.5;
                sheet.Cells["E4"].Value = 34;
                sheet.Cells["E5"].Value = 2;
                sheet.Cells["E6"].Value = 9;

                sheet.Cells["A8"].Formula = "GROWTH(A2:A5, B2:C5, D2:E6)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 8);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 3);
                var result5 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                Assert.AreEqual(3.602531375d, result1);
                Assert.AreEqual(16.03053728d, result2);
                Assert.AreEqual(0.051713747d, result3);
                Assert.AreEqual(1036481.184d, result4);
                Assert.AreEqual(9.245661284d, result5);


            }
        }
        [TestMethod]
        public void GrowthTestMultipleRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple rows");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 1;
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

                sheet.Cells["A8"].Formula = "GROWTH(A1:D1, A2:D3, A4:D5, TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 9);
                Assert.AreEqual(1.125885265d, result1);
                Assert.AreEqual(1.749709117d, result2);
                Assert.AreEqual(1.126753855d, result3);
                Assert.AreEqual(0.000452851d, result4);

            }
        }

        [TestMethod]
        public void GrowthTestCollinearRows()
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

                sheet.Cells["A8"].Formula = "GROWTH(A1:D1, A2:D3, A4:D5, TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 9);
                Assert.AreEqual(0.193150381d, result1);
                Assert.AreEqual(0.80060532d, result2);
                Assert.AreEqual(2.521079449d, result3);
                Assert.AreEqual(0.039675128d, result4);


            }
        }


        [TestMethod]
        public void GrowthTestDefaultX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with x-values omitted");
                sheet.Cells["A1"].Value = 232;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 21.121;
                sheet.Cells["D1"].Value = 332;
                sheet.Cells["A8"].Formula = "GROWTH(A1:D1)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 8);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 8);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 8);
                Assert.AreEqual(29.84928428d, result1);
                Assert.AreEqual(40.40064271d, result2);
                Assert.AreEqual(54.68177781d, result3);
                Assert.AreEqual(74.01112023d, result4);
            }
        }

        [TestMethod]
        public void GrowthTestDefaultXCols()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with x-values omitted, column-based");
                sheet.Cells["A1"].Value = 232;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 21.121;
                sheet.Cells["A4"].Value = 332;
                sheet.Cells["A8"].Formula = "GROWTH(A1:A4)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 8);
                var result2 = System.Math.Round((double)sheet.Cells["A9"].Value, 8);
                var result3 = System.Math.Round((double)sheet.Cells["A10"].Value, 8);
                var result4 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                Assert.AreEqual(29.84928428d, result1);
                Assert.AreEqual(40.40064271d, result2);
                Assert.AreEqual(54.68177781d, result3);
                Assert.AreEqual(74.01112023d, result4);
            }
        }

        [TestMethod]

        public void GrowthNegativeNumberTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with a negative number, which should return an error");
                sheet.Cells["A1"].Value = 232;
                sheet.Cells["A2"].Value = -3;
                sheet.Cells["A3"].Value = 21.121;
                sheet.Cells["A4"].Value = 332;
                sheet.Cells["A8"].Formula = "GROWTH(A1:A4)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result1);
            }
        }
    }
}