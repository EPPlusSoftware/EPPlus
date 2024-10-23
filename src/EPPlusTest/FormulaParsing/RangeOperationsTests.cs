using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class RangeOperationsTests : TestBase
    {
        [TestMethod]
        public void IntersectOperatorWithMultipleRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Formula = "SUM(A1:A3 A2:A4 OFFSET(A1, 1, 0))";
                sheet.Calculate();
                var result = sheet.Cells["A5"].Value;
                Assert.AreEqual(2d, result);
            }
        }

        [TestMethod]
        public void AdditionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 + B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(6d, result);
                SaveWorkbook("range.xlsx",package);
            }
        }

        [TestMethod]
        public void SubtractionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 - B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(1d, result);
            }
        }

        [TestMethod]
        public void MultiplicationOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 * B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(12d, result);
            }
        }

        [TestMethod]
        public void MultiplicationOperatorShouldCalculateRangeAndSingleValueRight()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 + 1)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(7d, result);
            }
        }

        [TestMethod]
        public void MultiplicationOperatorShouldCalculateRangeAndSingleValueLeft()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B3"].Formula = "SUM(1 - A1:A2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(-3d, result);
            }
        }

        [TestMethod]
        public void DivisionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "AVERAGE(A1:A2 / B1:B2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.75d, result);
            }
        }

        [TestMethod]
        public void ExpOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["B2"].Value = 4;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 ^ B1:B2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(85d, result);
            }
        }

        [TestMethod]
        public void EqualsOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Formula = "SUM(IF(A1:A3=3,1,2))";
                sheet.Calculate();
                var result = Math.Round((double)sheet.Cells["A4"].Value, 2);
                Assert.AreEqual(5d, result);
            }
        }

        [TestMethod]
        public void LessThanOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Formula = "SUM(IF(A1:A3<5,1,2))";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 2);
                Assert.AreEqual(4d, result);
            }
        }

        [TestMethod]
        public void LessThanOrEqualOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Formula = "SUM(IF(A1:A3<=3,1,2))";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 2);
                Assert.AreEqual(5d, result);
            }
        }

        [TestMethod]
        public void GreaterThanOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Formula = "SUM(IF(A1:A3>3,1,2))";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 2);
                Assert.AreEqual(4d, result);
            }
        }

        [TestMethod]
        public void GreaterThanOrEqualOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Formula = "SUM(IF(A1:A3>=4,1,2))";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 2);
                Assert.AreEqual(4d, result);
            }
        }

        [TestMethod]
        public void ConcatOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "c";
                sheet.Cells["B1"].Value = "b";
                sheet.Cells["B2"].Value = "d";
                sheet.Cells["B3"].Formula = "CONCAT(A1:A2 & B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value.ToString();
                Assert.AreEqual("abcd", result);
            }
        }

        [TestMethod]
        public void CachedRangeExpressionShouldCalculateTwice()
        {
            using (var p = new ExcelPackage())
            {
                var workbook = p.Workbook;
                var sheet = workbook.Worksheets.Add("ws");

                for (int i = 3; i < 9; i++)
                {
                    sheet.Cells[i, 2].Value = i - 2.0d;
                }

                sheet.Cells["C3"].Formula = "B3*2";
                sheet.Cells["C4:C8"].Formula = "B4*2";

                workbook.Calculate();

                Assert.AreEqual(1.0, sheet.Cells["B3"].Value);
                Assert.AreEqual(2.0, sheet.Cells["C3"].Value);
                Assert.AreEqual(2.0, sheet.Cells["B4"].Value);
                Assert.AreEqual(4.0, sheet.Cells["C4"].Value);

                sheet.Cells["B3"].Value = 300.0;
                sheet.Cells["B4"].Value = 300.0;

                workbook.Calculate();

                Assert.AreEqual(600.0, sheet.Cells["C3"].Value);
                Assert.AreEqual(600.0, sheet.Cells["C4"].Value);
            }
        }
    }
}
