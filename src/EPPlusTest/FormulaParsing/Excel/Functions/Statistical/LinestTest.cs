using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LinestTest
    {
        [TestMethod]
        public void LinestNoConstInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5, B2:B5,,FALSE)";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["A8"].Value);
                Assert.AreEqual(1d, sheet.Cells["B8"].Value);
            }
        }

        [TestMethod]
        public void LinestNoStatInput()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:B5,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                Assert.AreEqual(2.310344828d, result);
                Assert.AreEqual(0d, sheet.Cells["B8"].Value);
            }
        }

        [TestMethod]
        public void LinestTestWithStatsAndConst()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with statistics");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 1;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:B5,TRUE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 1);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 1);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 2);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 1);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 1);
                Assert.AreEqual(1.4d, result1);
                Assert.AreEqual(3.4d, result2);
                Assert.AreEqual(1.587450787d, result3);
                Assert.AreEqual(2.969848481d, result4);
                Assert.AreEqual(0.28d, result5);
                Assert.AreEqual(3.54964787d, result6);
                Assert.AreEqual(0.777777778d, result7);
                Assert.AreEqual(2d, result8);
                Assert.AreEqual(9.8d, result9);
                Assert.AreEqual(25.2, result10);
            }
        }

        [TestMethod]
        public void LinestTestWithStatConstIsFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with statistics and const = false");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 1;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:B5,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = sheet.Cells["B9"].Value;
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 7);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                Assert.AreEqual(2.857142857d, result1);
                Assert.AreEqual(0d, result2);
                Assert.AreEqual(0.996592835d, result3);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result4);
                Assert.AreEqual(0.732600733d, result5);
                Assert.AreEqual(3.728908943d, result6);
                Assert.AreEqual(8.219178082d, result7);
                Assert.AreEqual(3d, result8);
                Assert.AreEqual(114.2857143d, result9);
                Assert.AreEqual(41.71428571d, result10);
            }
        }

        [TestMethod]
        public void LinestTestWithNoXs()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with no Xs and only one argument: ");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["A6"].Value = 11;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Formula = "LINEST(A2:A7)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.371428571d, result1);
                Assert.AreEqual(4.533333333d, result2);
            }
        }

        [TestMethod]
        public void LinestTestWithNoXsAndConstIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with no Xs and const is true: ");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["A6"].Value = 11;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Formula = "LINEST(A2:A7,,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(0.371428571d, result1);
                Assert.AreEqual(4.533333333d, result2);
            }
        }

        [TestMethod]
        public void LinestTestWithNoXsAndConstTrueStatTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with no Xs and const is true and stat is true: ");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["A6"].Value = 11;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Formula = "LINEST(A2:A7,,TRUE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                Assert.AreEqual(0.371428571d, result1);
                Assert.AreEqual(4.533333333d, result2);
                Assert.AreEqual(1.031081593d, result3);
                Assert.AreEqual(4.015485896d, result4);
                Assert.AreEqual(0.031422374d, result5);
                Assert.AreEqual(4.313323765d, result6);
                Assert.AreEqual(0.129767085d, result7);
                Assert.AreEqual(4d, result8);
                Assert.AreEqual(2.414285714d, result9);
                Assert.AreEqual(74.41904762d, result10);
            }
        }

        [TestMethod]
        public void LinestTestWithNoXsAndConstFalseStatTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with no Xs and const is false and stat is true: ");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["A6"].Value = 11;
                sheet.Cells["A7"].Value = 2;
                sheet.Cells["A8"].Formula = "LINEST(A2:A7,,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = sheet.Cells["B9"].Value;
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 7);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                Assert.AreEqual(1.417582418d, result1);
                Assert.AreEqual(0d, result2);
                Assert.AreEqual(0.464407618d, result3);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result4);
                Assert.AreEqual(0.65077627d, result5);
                Assert.AreEqual(4.43016632d, result6);
                Assert.AreEqual(9.317469205d, result7);
                Assert.AreEqual(5d, result8);
                Assert.AreEqual(182.8681319d, result9);
                Assert.AreEqual(98.13186813d, result10);
            }
        }

        [TestMethod]
        public void LinestTestUnevenSizes()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test where datapoints are equal but size is not");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 1;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["C3"].Value = 7;
                sheet.Cells["C4"].Value = 2;
                sheet.Cells["C5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:B3,C2:C5,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void LinestMultipleXRangesSeveralColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["C2"].Value = 6;
                sheet.Cells["C3"].Value = 5;
                sheet.Cells["C4"].Value = 2;
                sheet.Cells["C5"].Value = 0;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:C5)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 0);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 0);
                Assert.AreEqual(0d, result1);
                Assert.AreEqual(2d, result2);
                Assert.AreEqual(1d, result3);
            }
        }

        [TestMethod]
        public void LinestMultipleXRangesSeveralRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["D2"].Value = 7;
                sheet.Cells["E2"].Value = 0;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["C3"].Value = 3;
                sheet.Cells["D3"].Value = 6;
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["C4"].Value = 8;
                sheet.Cells["D4"].Value = 5;
                sheet.Cells["E4"].Value = 1;
                sheet.Cells["A8"].Formula = "LINEST(A2:E2,A3:E4)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                Assert.AreEqual(0.450151057d, result1);
                Assert.AreEqual(-0.854984894d, result2);
                Assert.AreEqual(6.19939577d, result3);
            }
        }

        [TestMethod]
        public void LinestMultipleXRangesTwoByTwo() //This test returns failed because of collinearity
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges");
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["A3"].Value = 12;
                sheet.Cells["B2"].Value = 34;
                sheet.Cells["B3"].Value = 65;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 7;
                sheet.Cells["A8"].Formula = "LINEST(A2:A3,B2:C3)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 0);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                Assert.AreEqual(0d, result1);
                Assert.AreEqual(0.096774194d, result2);
                Assert.AreEqual(5.709677419d, result3);
            }
        }

        [TestMethod]
        public void LinestMultipleRegressionWithStats()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges and stats");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["D2"].Value = 7;
                sheet.Cells["E2"].Value = 0;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["C3"].Value = 3;
                sheet.Cells["D3"].Value = 6;
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["C4"].Value = 8;
                sheet.Cells["D4"].Value = 5;
                sheet.Cells["E4"].Value = 1;
                sheet.Cells["A8"].Formula = "LINEST(A2:E2,A3:E4,TRUE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 8);
                var result4 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["C9"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result9 = sheet.Cells["C10"].Value;
                var result10 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result11 = (int)sheet.Cells["B11"].Value;
                var result12 = sheet.Cells["C11"].Value;
                var result13 = System.Math.Round((double)sheet.Cells["A12"].Value, 8);
                var result14 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                var result15 = sheet.Cells["C12"].Value;
                Assert.AreEqual(0.450151057d, result1);
                Assert.AreEqual(-0.854984894d, result2);
                Assert.AreEqual(6.19939577d, result3);
                Assert.AreEqual(0.81889275d, result4);
                Assert.AreEqual(1.492093601d, result5);
                Assert.AreEqual(7.119188135d, result6);
                Assert.AreEqual(0.250122479d, result7);
                Assert.AreEqual(4.711302858d, result8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result9);
                Assert.AreEqual(0.333551109d, result10);
                Assert.AreEqual(2, result11);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result12);
                Assert.AreEqual(14.80725076d, result13);
                Assert.AreEqual(44.39274924d, result14);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result15);
            }
        }

        [TestMethod]
        public void LinestMultipleRegressionWithStatsAndConstIsFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges and stats");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["D2"].Value = 7;
                sheet.Cells["E2"].Value = 0;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["C3"].Value = 3;
                sheet.Cells["D3"].Value = 6;
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["C4"].Value = 8;
                sheet.Cells["D4"].Value = 5;
                sheet.Cells["E4"].Value = 1;
                sheet.Cells["A8"].Formula = "LINEST(A2:E2,A3:E4,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 0);
                var result4 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result6 = sheet.Cells["C9"].Value;
                var result7 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result9 = sheet.Cells["C10"].Value;
                var result10 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result11 = (int)sheet.Cells["B11"].Value;
                var result12 = sheet.Cells["C11"].Value;
                var result13 = System.Math.Round((double)sheet.Cells["A12"].Value, 8);
                var result14 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                var result15 = sheet.Cells["C12"].Value;
                Assert.AreEqual(0.778248214d, result1);
                Assert.AreEqual(0.263826409d, result2);
                Assert.AreEqual(0d, result3);
                Assert.AreEqual(0.697161675d, result4);
                Assert.AreEqual(0.727487085d, result5);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result6);
                Assert.AreEqual(0.607537607d, result7);
                Assert.AreEqual(4.517526365d, result8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result9);
                Assert.AreEqual(2.32202225d, result10);
                Assert.AreEqual(3, result11);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result12);
                Assert.AreEqual(94.77586663d, result13);
                Assert.AreEqual(61.22413337d, result14);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result15);
            }
        }

        [TestMethod]
        public void LinestCollinearityTest() //This test returns failed because of collinearity
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with redundant x-sets");
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["A3"].Value = 12;
                sheet.Cells["B2"].Value = 34;
                sheet.Cells["B3"].Value = 65;
                sheet.Cells["C2"].Value = 8;
                sheet.Cells["C3"].Value = 7;
                sheet.Cells["D2"].Value = 15;
                sheet.Cells["D3"].Value = 2431;
                sheet.Cells["E2"].Value = 2534;
                //sheet.Cells["E3"].Value = 6769;
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["A8"].Formula = "LINEST(A2:A3,B2:E3,false,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 0);
                var result4 = System.Math.Round((double)sheet.Cells["D8"].Value, 0);
                var result5 = System.Math.Round((double)sheet.Cells["E8"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["A9"].Value, 0);
                var result7 = System.Math.Round((double)sheet.Cells["B9"].Value, 0);
                var result8 = System.Math.Round((double)sheet.Cells["C9"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["D9"].Value, 0);
                var result10 = System.Math.Round((double)sheet.Cells["E9"].Value, 0);
                Assert.AreEqual(0.000708383d, result1);
                Assert.AreEqual(0d, result2);
                Assert.AreEqual(0d, result3);
                Assert.AreEqual(0d, result4);
                Assert.AreEqual(7.204958678d, result5);
                Assert.AreEqual(0d, result6);
                Assert.AreEqual(0d, result7);
                Assert.AreEqual(0d, result8);
                Assert.AreEqual(0d, result9);
                Assert.AreEqual(0d, result10);
            }
        }

        [TestMethod]
        public void LinestRemovalOfRedundantVariablesTest() //This test returns failed because of collinearity
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with where independent values are completely dependent of one another");
                sheet.Cells["A2"].Value = 10;
                sheet.Cells["A3"].Value = 20;
                sheet.Cells["A4"].Value = 30;
                sheet.Cells["A5"].Value = 40;
                sheet.Cells["A6"].Value = 50;
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 8;
                sheet.Cells["B5"].Value = 7;
                sheet.Cells["B6"].Value = 9;
                sheet.Cells["C2"].Value = 11;
                sheet.Cells["C3"].Value = 20;
                sheet.Cells["C4"].Value = 32;
                sheet.Cells["C5"].Value = 29;
                sheet.Cells["C6"].Value = 35;
                sheet.Cells["A8"].Formula = "LINEST(A2:A6,B2:C6,TRUE,true)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                var result2 = sheet.Cells["B8"].Value;
                var result3 = sheet.Cells["C8"].Value;
                Assert.AreEqual(0d, result1);
                Assert.AreEqual(0.096774194d, result2);
                Assert.AreEqual(5.709677419d, result3);
            }
        }

        [TestMethod]
        public void LinestWithMultipleXNoCollinearity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with multiple x-variables, but there is no collinearity: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 9;
                sheet.Cells["A3"].Value = 18;
                sheet.Cells["A4"].Value = 0;
                sheet.Cells["A5"].Value = 4;
                sheet.Cells["A6"].Value = 45;
                sheet.Cells["B1"].Value = 50;
                sheet.Cells["B2"].Value = 345;
                sheet.Cells["B3"].Value = 3.4983;
                sheet.Cells["B4"].Value = 234;
                sheet.Cells["B5"].Value = 876;
                sheet.Cells["B6"].Value = 876;
                sheet.Cells["C1"].Value = 2738;
                sheet.Cells["C2"].Value = 29810;
                sheet.Cells["C3"].Value = 4309;
                sheet.Cells["C4"].Value = 95;
                sheet.Cells["C5"].Value = 34.0678;
                sheet.Cells["C6"].Value = 561.4823;
                sheet.Cells["D1"].Value = 2;
                sheet.Cells["D2"].Value = 8;
                sheet.Cells["D3"].Value = 6666;
                sheet.Cells["D4"].Value = 5;
                sheet.Cells["D5"].Value = 544.45;
                sheet.Cells["D6"].Value = 7654;
                sheet.Cells["E1"].Value = 543;
                sheet.Cells["E2"].Value = 890;
                sheet.Cells["E3"].Value = 876;
                sheet.Cells["E4"].Value = 8765;
                sheet.Cells["E5"].Value = 3487.298;
                sheet.Cells["E6"].Value = 32.1;
                sheet.Cells["F1"].Value = 50;
                sheet.Cells["F2"].Value = 30;
                sheet.Cells["F3"].Value = 2397.346;
                sheet.Cells["F4"].Value = 423.789;
                sheet.Cells["F5"].Value = 432.4;
                sheet.Cells["F6"].Value = 746;
                sheet.Cells["F10"].Formula = "LINEST(A1:A6, B1:F6, FALSE, TRUE)";
                sheet.Calculate();

                //When debugging, this test returns the same as excel.

            }
        }

    }
}
