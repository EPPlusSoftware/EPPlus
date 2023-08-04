using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class FDotInvTest
    {
        [TestMethod]
        public void FDotInvTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(0.01, 6, 4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.10930991d, result);

            }
        }

        [TestMethod]
        public void FDotInvTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(0.08, 6, 9)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 6);
                Assert.AreEqual(0.303175d, result);

            }
        }

        [TestMethod]
        public void FDotInvShouldReturnNumWrongProbability()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(1.02, 6, 9)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void FDotInvShouldReturnNumWrongDF1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(1.02, 0.2, 9)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void FDotInvShouldReturnNumWrongDF2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(1.02, 4, 0.9)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void FDotInvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV(0.03, 6, 2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 6);
                Assert.AreEqual(0.150265d, result);

            }
        }
    }
}
