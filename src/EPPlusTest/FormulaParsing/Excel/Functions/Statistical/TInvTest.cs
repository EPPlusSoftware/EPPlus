using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TInvTest : TestBase
    {
        [TestMethod]
        public void TInvCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV(0.75,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(0.8164966, result);
            }
        }

        [TestMethod]
        public void TInvDegreesOfFreedomLessThanOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM! if df is less than 1: ");
                sheet.Cells["A1"].Formula = "T.INV(0.45,0.99)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TInvIncorrectProbabilityOver()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM!");
                sheet.Cells["A1"].Formula = "T.INV(1.06,3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TInvIncorrectProbabilityUnder()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM!");
                sheet.Cells["A1"].Formula = "T.INV(-0.5,3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
        [TestMethod]
        public void TInvCorrectReuslt2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV(0.99,18)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(2.5523796, result);
            }
        }
        [TestMethod]
        public void TInvCorrectReuslt3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV(0.12,9.678)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(-1.25814181, result);
            }
        }

        [TestMethod]
        public void TInvCorrectReuslt4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV(0.86,34)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(1.09781350, result);
            }
        }
    }
}
