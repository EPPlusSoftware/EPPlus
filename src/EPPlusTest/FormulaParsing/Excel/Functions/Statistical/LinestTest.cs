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
                sheet.Cells["A8"].Formula = "LINEST(A2:A7,,FALSE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 0);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 9);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 9);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 8);
                Assert.AreEqual(1.417582418d, result1);
                Assert.AreEqual(0d, result2);
                Assert.AreEqual(0.464407618d, result3);
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
        public void LinestMultipleXRanges()
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
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 9);
                var result3 = System.Math.Round((double)sheet.Cells["C8"].Value, 9);
                Assert.AreEqual(0.007556081d, result1);
                Assert.AreEqual(0.325533857d, result2);
                Assert.AreEqual(3.490035743d, result3);
            }
        }
    }
}
