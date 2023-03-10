using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
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
    }
}
