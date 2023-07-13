using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TTestTest : TestBase
    {

        [TestMethod]
        public void TTestWithTypeEqualsPaired()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("T-test with type = 1 as input: ");
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
                sheet.Cells["B10"].Formula = "T.TEST(A1:A9,B1:B9,1,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                Assert.AreEqual(0.098007892d, result);
            }
        }

        [TestMethod]
        public void TTestWithEqualVariance()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with type = 2 as input: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 54;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["A7"].Value = 9;
                sheet.Cells["A8"].Value = 3;
                sheet.Cells["A9"].Value = 789;
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 3;
                sheet.Cells["B3"].Value = 8;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["B5"].Value = 43;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A9,B1:B6,1,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                Assert.AreEqual(0.218530672d, result);
            }
        }

        [TestMethod]
        public void TTestWithUnequalVariance()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with type = 3 as input: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 54;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["A7"].Value = 9;
                sheet.Cells["A8"].Value = 3;
                sheet.Cells["A9"].Value = 789;
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 3;
                sheet.Cells["B3"].Value = 8;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["B5"].Value = 43;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A9,B1:B6,1,3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 8);
                Assert.AreEqual(0.17474296d, result);
            }
        }

        [TestMethod]
        public void TTestType2WithEqualDataPoints()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with type = 2 as input and equal samples: ");
                sheet.Cells["A1"].Value = 19;
                sheet.Cells["A2"].Value = 456;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 8432;
                sheet.Cells["A5"].Value = 0;
                sheet.Cells["B1"].Value = 12;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 789;
                sheet.Cells["B5"].Value = 5;
                sheet.Cells["B6"].Formula = "T.TEST(A1:A5,B1:B5,2,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B6"].Value, 9);
                Assert.AreEqual(0.361518411d, result);
            }
        }

        [TestMethod]
        public void TTestType3WithEqualDataPoints()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with type = 2 as input and equal samples: ");
                sheet.Cells["A1"].Value = 19;
                sheet.Cells["A2"].Value = 456;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 1;
                sheet.Cells["A5"].Value = 0;
                sheet.Cells["B1"].Value = 12;
                sheet.Cells["B2"].Value = 9;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 789;
                sheet.Cells["B5"].Value = 5;
                sheet.Cells["B6"].Formula = "T.TEST(A1:A5,B1:B5,2,3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B6"].Value, 9);
                Assert.AreEqual(0.718549187d, result);
            }
        }

        [TestMethod]
        public void TTestType3EqualDataPoints()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Equal datapoints: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 9;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["B6"].Formula = "T.TEST(A1:A3,A4:A6,2,3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B6"].Value, 8);
                Assert.AreEqual(0.25194841d, result);
            }
        }

        [TestMethod]
        public void TTestType2EqualDataPoints()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Equal datapoints: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 9;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["B6"].Formula = "T.TEST(A1:A3,A4:A6,2,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B6"].Value, 7);
                Assert.AreEqual(0.2161194d, result);
            }
        }

        [TestMethod]

        public void TTestType1MessyArrays()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Equal datapoints: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = "HKOLWD";
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = "";
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 9;
                sheet.Cells["B4"].Value = "AS";
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A6,B1:B6,2,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                Assert.AreEqual(0.422649731d, result);
            }
        }

        [TestMethod]
        public void TTestType2MessyArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Equal datapoints: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = "HKOLWD";
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = "";
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 9;
                sheet.Cells["B4"].Value = "AS";
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A6,B1:B6,2,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                Assert.AreEqual(0.607797102d, result);
            }
        }

        [TestMethod]
        public void TTestType3MessyArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Equal datapoints: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = "HKOLWD";
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = "";
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 9;
                sheet.Cells["B4"].Value = "AS";
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A6,B1:B6,2,3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B7"].Value, 9);
                Assert.AreEqual(0.616006838d, result);
            }
        }

        [TestMethod]
        public void TTestDifferentDataPointsType1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with type = 1 as input, not equal sample sizes: ");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 54;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["A7"].Value = 9;
                sheet.Cells["A8"].Value = 3;
                sheet.Cells["A9"].Value = 789;
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 3;
                sheet.Cells["B3"].Value = 8;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["B5"].Value = 43;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A9,B1:B6,1,1)";
                sheet.Calculate();
                var result = sheet.Cells["B7"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);
            }
        }

        [TestMethod]
        public void TTestIncorrectTail()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test incorrect tail");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 54;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["A7"].Value = 9;
                sheet.Cells["A8"].Value = 3;
                sheet.Cells["A9"].Value = 789;
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 3;
                sheet.Cells["B3"].Value = 8;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["B5"].Value = 43;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A9,B1:B6,5,2)";
                sheet.Calculate();
                var result = sheet.Cells["B7"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TTestIncorrectType()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect type");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 8;
                sheet.Cells["A5"].Value = 54;
                sheet.Cells["A6"].Value = 12;
                sheet.Cells["A7"].Value = 9;
                sheet.Cells["A8"].Value = 3;
                sheet.Cells["A9"].Value = 789;
                sheet.Cells["B1"].Value = 5;
                sheet.Cells["B2"].Value = 3;
                sheet.Cells["B3"].Value = 8;
                sheet.Cells["B4"].Value = 3;
                sheet.Cells["B5"].Value = 43;
                sheet.Cells["B6"].Value = 12;
                sheet.Cells["B7"].Formula = "T.TEST(A1:A9,B1:B6,1,4)";
                sheet.Calculate();
                var result = sheet.Cells["B7"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TTestDivisonByZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with empty sets");
                sheet.Cells["A1"].Value = "";
                sheet.Cells["A2"].Value = "";
                sheet.Cells["A3"].Value = "";
                sheet.Cells["A4"].Value = "";
                sheet.Cells["A5"].Value = "";
                sheet.Cells["A6"].Value = "";
                sheet.Cells["B1"].Value = "";
                sheet.Cells["B2"].Value = "";
                sheet.Cells["B3"].Value = "";
                sheet.Cells["B4"].Value = "";
                sheet.Cells["B5"].Value = "";
                sheet.Cells["B6"].Value = "";
                sheet.Cells["B7"].Formula = "T.TEST(A1:A6,B1:B6,1,2)";
                sheet.Calculate();
                var result = sheet.Cells["B7"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

        [TestMethod]
        public void TTestOneSampleBelowTwo()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with empty sets");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 4;
                sheet.Cells["A3"].Value = 8;
                sheet.Cells["A4"].Value = 6;
                sheet.Cells["A5"].Value = 4;
                sheet.Cells["A6"].Value = 5;
                sheet.Cells["B1"].Value = 9;
                sheet.Cells["B2"].Value = "fds";
                sheet.Cells["B3"].Value = "";
                sheet.Cells["B4"].Value = "";
                sheet.Cells["B5"].Value = "j";
                sheet.Cells["B6"].Value = "";
                sheet.Cells["B7"].Formula = "T.TEST(A1:A6,B1:B6,1,2)";
                sheet.Calculate();
                var result = sheet.Cells["B7"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

    }
}
