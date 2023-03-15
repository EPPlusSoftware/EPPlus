using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.ArrayTests
{
    [TestClass]
    public class MathFunctionsArrayTests
    {
        [TestMethod]
        public void AbsShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ABS(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void SignShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("SIGN(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(1d, sheet.Cells["B2"].Value);
                Assert.AreEqual(1d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void PowerShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("POWER(A1:A3,2)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(4d, sheet.Cells["B2"].Value);
                Assert.AreEqual(9d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void SqrtShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 9;
                sheet.Cells["A2"].Value = 16;
                sheet.Cells["A3"].Value = 25;
                sheet.Cells["B1:B3"].CreateArrayFormula("SQRT(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(3d, sheet.Cells["B1"].Value);
                Assert.AreEqual(4d, sheet.Cells["B2"].Value);
                Assert.AreEqual(5d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void CeilingShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9;
                sheet.Cells["A2"].Value = 2.9;
                sheet.Cells["A3"].Value = 3.3;
                sheet.Cells["B1:B3"].CreateArrayFormula("CEILING(A1:A3,1)");
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["B1"].Value);
                Assert.AreEqual(3d, sheet.Cells["B2"].Value);
                Assert.AreEqual(4d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsoCeilingShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9;
                sheet.Cells["A2"].Value = 2.9;
                sheet.Cells["A3"].Value = 3.3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ISO.CEILING(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["B1"].Value);
                Assert.AreEqual(3d, sheet.Cells["B2"].Value);
                Assert.AreEqual(4d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void CeilingPreciseShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9;
                sheet.Cells["A2"].Value = 2.9;
                sheet.Cells["A3"].Value = 3.3;
                sheet.Cells["B1:B3"].CreateArrayFormula("CEILING.PRECISE(A1:A3,2)");
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["B1"].Value);
                Assert.AreEqual(4d, sheet.Cells["B2"].Value);
                Assert.AreEqual(4d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void CeilingMathShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9;
                sheet.Cells["A2"].Value = 2.9;
                sheet.Cells["A3"].Value = 3.3;
                sheet.Cells["B1:B3"].CreateArrayFormula("CEILING.MATH(A1:A3,1)");
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["B1"].Value);
                Assert.AreEqual(3d, sheet.Cells["B2"].Value);
                Assert.AreEqual(4d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void EvenShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("EVEN(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["B1"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.AreEqual(4, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void OddShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ODD(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["B1"].Value);
                Assert.AreEqual(3, sheet.Cells["B2"].Value);
                Assert.AreEqual(3, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void RoundShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("ROUND(A1:A3,3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 3);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 3);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 3);
                Assert.AreEqual(1.914, v1);
                Assert.AreEqual(2.6, v2);
                Assert.AreEqual(3.102, v3);
            }
        }

        [TestMethod]
        public void RoundDownShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("ROUNDDOWN(A1:A3,2)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.91, v1);
                Assert.AreEqual(2.59, v2);
                Assert.AreEqual(3.1, v3);
            }
        }

        [TestMethod]
        public void RoundUpShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("ROUNDUP(A1:A3,3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 3);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 3);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 3);
                Assert.AreEqual(1.914, v1);
                Assert.AreEqual(2.6, v2);
                Assert.AreEqual(3.102, v3);
            }
        }

        [TestMethod]
        public void TruncShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("TRUNC(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void DegreesShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("DEGREES(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(57.2958, v1);
                Assert.AreEqual(114.5916, v2);
                Assert.AreEqual(171.8873, v3);
            }
        }

        [TestMethod]
        public void RadiansShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("RADIANS(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(0.0175, v1);
                Assert.AreEqual(0.0349, v2);
                Assert.AreEqual(0.0524, v3);
            }
        }

        [TestMethod]
        public void CosShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("COS(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(0.5403, v1);
                Assert.AreEqual(-0.4161, v2);
                Assert.AreEqual(-0.99, v3);
            }
        }

        [TestMethod]
        public void AcosShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = -0.5;
                sheet.Cells["A2"].Value = 0.1;
                sheet.Cells["A3"].Value = 0.9;
                sheet.Cells["B1:B3"].CreateArrayFormula("ACOS(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(2.0944, v1);
                Assert.AreEqual(1.4706, v2);
                Assert.AreEqual(0.451, v3);
            }
        }

        [TestMethod]
        public void AcoshShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ACOSH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(0d, v1);
                Assert.AreEqual(1.317, v2);
                Assert.AreEqual(1.7627, v3);
            }
        }

        [TestMethod]
        public void SecShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("SEC(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(1.8508d, v1);
                Assert.AreEqual(-2.403, v2);
                Assert.AreEqual(-1.0101, v3);
            }
        }

        [TestMethod]
        public void SechShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("SECH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(0.6481d, v1);
                Assert.AreEqual(0.2658, v2);
                Assert.AreEqual(0.0993, v3);
            }
        }

        [TestMethod]
        public void SinShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("SIN(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(0.8415, v1);
                Assert.AreEqual(0.9093, v2);
                Assert.AreEqual(0.1411, v3);
            }
        }

        [TestMethod]
        public void AsinShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = -0.5;
                sheet.Cells["A2"].Value = 0.1;
                sheet.Cells["A3"].Value = 0.9;
                sheet.Cells["B1:B3"].CreateArrayFormula("ASIN(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 4);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 4);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 4);
                Assert.AreEqual(-0.5236, v1);
                Assert.AreEqual(0.1002, v2);
                Assert.AreEqual(1.1198, v3);
            }
        }

        [TestMethod]
        public void AcotShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ACOT(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.79, v1);
                Assert.AreEqual(0.46, v2);
                Assert.AreEqual(0.32, v3);
            }
        }

        [TestMethod]
        public void AcothShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("ACOTH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.58, v1);
                Assert.AreEqual(0.41, v2);
                Assert.AreEqual(0.33, v3);
            }
        }

        [TestMethod]
        public void AsinhShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ASINH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.88, v1);
                Assert.AreEqual(1.44, v2);
                Assert.AreEqual(1.82, v3);
            }
        }

        [TestMethod]
        public void AtanShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ATAN(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.79, v1);
                Assert.AreEqual(1.11, v2);
                Assert.AreEqual(1.25, v3);
            }
        }

        [TestMethod]
        public void AtanhShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("S47heet1");

                sheet.Cells["A1"].Value = -0.5;
                sheet.Cells["A2"].Value = 0.1;
                sheet.Cells["A3"].Value = 0.9;
                sheet.Cells["B1:B3"].CreateArrayFormula("ATANH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(-0.55, v1);
                Assert.AreEqual(0.1, v2);
                Assert.AreEqual(1.47, v3);
            }
        }

        [TestMethod]
        public void CotShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("COT(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.64, v1);
                Assert.AreEqual(-0.46, v2);
                Assert.AreEqual(-7.02, v3);
            }
        }

        [TestMethod]
        public void CothShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("COTH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.31, v1);
                Assert.AreEqual(1.04, v2);
                Assert.AreEqual(1.0, v3);
            }
        }

        [TestMethod]
        public void CscShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("CSC(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.19, v1);
                Assert.AreEqual(1.1, v2);
                Assert.AreEqual(7.09, v3);
            }
        }

        [TestMethod]
        public void CschShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("CSCH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.85, v1);
                Assert.AreEqual(0.28, v2);
                Assert.AreEqual(0.1, v3);
            }
        }

        [TestMethod]
        public void ExpShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("EXP(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(2.72, v1);
                Assert.AreEqual(7.39, v2);
                Assert.AreEqual(20.09, v3);
            }
        }

        [TestMethod]
        public void FloorShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("FLOOR(A1:A3,1)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void FloorPreciseShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("FLOOR.PRECISE(A1:A3,1)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void FloorMathShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1.9135;
                sheet.Cells["A2"].Value = 2.5999;
                sheet.Cells["A3"].Value = 3.1015;
                sheet.Cells["B1:B3"].CreateArrayFormula("FLOOR.MATH(A1:A3,1)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void LnShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("LN(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0d, v1);
                Assert.AreEqual(0.69, v2);
                Assert.AreEqual(1.1, v3);
            }
        }

        [TestMethod]
        public void LogShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("LOG(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0d, v1);
                Assert.AreEqual(0.3, v2);
                Assert.AreEqual(0.48, v3);
            }
        }

        [TestMethod]
        public void Log10ShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("LOG10(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0d, v1);
                Assert.AreEqual(0.3, v2);
                Assert.AreEqual(0.48, v3);
            }
        }

        [TestMethod]
        public void ModShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("MOD(A1:A3,2)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1d, v1);
                Assert.AreEqual(0d, v2);
                Assert.AreEqual(1d, v3);
            }
        }

        [TestMethod]
        public void RomanShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ROMAN(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual("I", sheet.Cells["B1"].Value);
                Assert.AreEqual("II", sheet.Cells["B2"].Value);
                Assert.AreEqual("III", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void SinhShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("SINH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.18d, v1);
                Assert.AreEqual(3.63, v2);
                Assert.AreEqual(10.02, v3);
            }
        }

        [TestMethod]
        public void TanShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("TAN(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.56d, v1);
                Assert.AreEqual(-2.19, v2);
                Assert.AreEqual(-0.14, v3);
            }
        }

        [TestMethod]
        public void TanhShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("TANH(A1:A3)");
                sheet.Calculate();
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 2);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 2);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(0.76d, v1);
                Assert.AreEqual(0.96, v2);
                Assert.AreEqual(1d, v3);
            }
        }
    }
}
