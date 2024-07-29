using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LogestTest
    {
        [TestMethod]
        public void LogestNoConstInput()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A5, B2:B5,,FALSE)";
                sheet.Calculate();
                Assert.AreEqual(1.751115956d, System.Math.Round((double)sheet.Cells["A8"].Value, 9));
                Assert.AreEqual(1.194315591d, System.Math.Round((double)sheet.Cells["B8"].Value, 9));
            }
        }

        [TestMethod]
        public void LogestNoStatInput()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A5,B2:B5,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                Assert.AreEqual(1.850326716d, result);
                Assert.AreEqual(1d, sheet.Cells["B8"].Value);
            }
        }

        [TestMethod]
        public void LogestTestWithStatsAndConst()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A5,B2:B5,TRUE,TRUE)";
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
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                Assert.AreEqual(1.690449345d, result1);
                Assert.AreEqual(1.916789388d, result2);
                Assert.AreEqual(0.394148965d, result3);
                Assert.AreEqual(0.737385194d, result4);
                Assert.AreEqual(0.470078317d, result5);
                Assert.AreEqual(0.88134388d, result6);
                Assert.AreEqual(1.7741426d, result7);
                Assert.AreEqual(2d, result8);
                Assert.AreEqual(1.378095486d, result9);
                Assert.AreEqual(1.553534068d, result10);
            }
        }

        [TestMethod]
        public void LogestTestWithStatConstIsFalse()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A5,B2:B5,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = sheet.Cells["B9"].Value;
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                Assert.AreEqual(2.234114741d, result1);
                Assert.AreEqual(1d, result2);
                Assert.AreEqual(0.226690276d, result3);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result4);
                Assert.AreEqual(0.807373214d, result5);
                Assert.AreEqual(0.848197344d, result6);
                Assert.AreEqual(12.574158032d, result7);
                Assert.AreEqual(3d, result8);
                Assert.AreEqual(9.046336342d, result9);
                Assert.AreEqual(2.158316204d, result10);
            }
        }

        [TestMethod]
        public void LogestTestWithNoXs()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A7)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(1.134094872d, result1);
                Assert.AreEqual(2.810925606d, result2);
            }
        }

        [TestMethod]
        public void LogestTestWithNoXsAndConstIsTrue()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A7,,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                Assert.AreEqual(1.134094872d, result1);
                Assert.AreEqual(2.810925606d, result2);
            }
        }

        [TestMethod]
        public void LogestTestWithNoXsAndConstTrueStatTrue()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A7,,TRUE,TRUE)";
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
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                Assert.AreEqual(1.134094872d, result1);
                Assert.AreEqual(2.810925606d, result2);
                Assert.AreEqual(0.242692744d, result3);
                Assert.AreEqual(0.945152447d, result4);
                Assert.AreEqual(0.062976548d, result5);
                Assert.AreEqual(1.015256588d, result6);
                Assert.AreEqual(0.268836592d, result7);
                Assert.AreEqual(4d, result8);
                Assert.AreEqual(0.277102226d, result9);
                Assert.AreEqual(4.122983757d, result10);
            }
        }

        [TestMethod]
        public void LogestTestWithNoXsAndConstFalseStatTrue()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A7,,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = sheet.Cells["B9"].Value;
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 8);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                Assert.AreEqual(1.439560782d, result1);
                Assert.AreEqual(1d, result2);
                Assert.AreEqual(0.108490801d, result3);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result4);
                Assert.AreEqual(0.692832622d, result5);
                Assert.AreEqual(1.034936276d, result6);
                Assert.AreEqual(11.27777021d, result7);
                Assert.AreEqual(5d, result8);
                Assert.AreEqual(12.07954182d, result9);
                Assert.AreEqual(5.355465482d, result10);
            }
        }

        [TestMethod]
        public void LogestTestUnevenSizes()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:B3,C2:C5,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), result1);

            }
        }

        [TestMethod]
        public void LogestMultipleXRangesSeveralColumns()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A5,B2:C5)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                Assert.AreEqual(0.923986526d, result1);
                Assert.AreEqual(1.669991606d, result2);
                Assert.AreEqual(1.71813453d, result3);
            }
        }

        [TestMethod]
        public void LogestMultipleXRangesSeveralRows()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with multiple x-ranges");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["C2"].Value = 5;
                sheet.Cells["D2"].Value = 7;
                sheet.Cells["E2"].Value = 1;
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:E2,A3:E4, FALSE, TRUE)";
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
                var result11 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result12 = sheet.Cells["C11"].Value;
                var result13 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                var result14 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                var result15 = sheet.Cells["C12"].Value;
                Assert.AreEqual(1.284511868d, result1);
                Assert.AreEqual(1.035289866d, result2);
                Assert.AreEqual(1d, result3);
                Assert.AreEqual(0.171842341d, result4);
                Assert.AreEqual(0.179317206d, result5);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result6);
                Assert.AreEqual(0.668015656d, result7);
                Assert.AreEqual(1.113518331d, result8);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result9);
                Assert.AreEqual(3.018285361d, result10);
                Assert.AreEqual(3d, result11);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result12);
                Assert.AreEqual(7.484883324d, result13);
                Assert.AreEqual(3.719769221d, result14);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result15);
            }
        }

        [TestMethod]
        public void LogestMultipleXRangesTwoByTwo()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A3,B2:C3)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 0);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                Assert.AreEqual(1d, result1);
                Assert.AreEqual(1.00932326d, result2);
                Assert.AreEqual(6.564670423d, result3);
            }
        }

        [TestMethod]
        public void LogestMultipleRegressionWithStats()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:E2,A3:E4,TRUE,TRUE)";
                sheet.Calculate();
                var result1 = sheet.Cells["A8"].Value;
                var result15 = sheet.Cells["C12"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result1);
            }
        }

        [TestMethod]
        public void LogestCollinearityTest()
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
                sheet.Cells["E3"].Value = 6769;
                sheet.Cells["A8"].Formula = "LOGEST(A2:A3,B2:E3,TRUE,TRUE)";
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
                Assert.AreEqual(1.000067932d, result1);
                Assert.AreEqual(1d, result2);
                Assert.AreEqual(1d, result3);
                Assert.AreEqual(1d, result4);
                Assert.AreEqual(7.576799201d, result5);
                Assert.AreEqual(0d, result6);
                Assert.AreEqual(0d, result7);
                Assert.AreEqual(0d, result8);
                Assert.AreEqual(0d, result9);
                Assert.AreEqual(0d, result10);
            }
        }

        [TestMethod]
        public void LinestRemovalOfRedundantVariablesTest()
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
                sheet.Cells["A8"].Formula = "LOGEST(A2:A6,B2:C6,TRUE,true)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["B9"].Value, 0);
                var result6 = System.Math.Round((double)sheet.Cells["C9"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result9 = sheet.Cells["C10"].Value;
                var result10 = System.Math.Round((double)sheet.Cells["A11"].Value, 8);
                var result11 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result12 = sheet.Cells["C11"].Value;
                var result13 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                var result14 = System.Math.Round((double)sheet.Cells["B12"].Value, 9);
                var result15 = sheet.Cells["C12"].Value;
                Assert.AreEqual(1.064146621d, result1);
                Assert.AreEqual(1d, result2);
                Assert.AreEqual(5.370304442d, result3);
                Assert.AreEqual(0.010462504d, result4);
                Assert.AreEqual(0d, result5);
                Assert.AreEqual(0.28116703d, result6);
                Assert.AreEqual(0.921697643d, result7);
                Assert.AreEqual(0.205342474d, result8);
                Assert.AreEqual(35.31302299d, result10);
                Assert.AreEqual(3d, result11);
                Assert.AreEqual(1.488992392d, result13);
                Assert.AreEqual(0.126496595d, result14);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result9);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result12);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result15);
            }
        }

        [TestMethod]
        public void LogestConstFalseZerosTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Const False and Collinearity test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["B2"].Value = 4;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["C2"].Value = 6;
                sheet.Cells["D1"].Value = 7;
                sheet.Cells["D2"].Value = 8;
                sheet.Cells["E1"].Value = 534;
                sheet.Cells["E2"].Value = 25464;
                sheet.Cells["D6"].Formula = "LOGEST(A1:A2, B1:E2, FALSE, FALSE)";
                sheet.Calculate();

                var result1 = System.Math.Round((double)sheet.Cells["D6"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["E6"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["F6"].Value, 0);
                var result4 = System.Math.Round((double)sheet.Cells["G6"].Value, 0);
                var result5 = System.Math.Round((double)sheet.Cells["H6"].Value, 0);
                Assert.AreEqual(1.000027889d, result1);
                Assert.AreEqual(0.997874723d, result2);
                Assert.AreEqual(1d, result3);
                Assert.AreEqual(1d, result4);
                Assert.AreEqual(1d, result5);
            }
        }
    }
}
