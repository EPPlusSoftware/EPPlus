using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TInv2tTest : TestBase
    {

        [TestMethod]
        public void TInv2tCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV.2T(0.75,2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.365148372, result);
            }
        }

        [TestMethod]
        public void TInv2tDegreesOfFreedomLessThanOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM! if df is less than 1: ");
                sheet.Cells["A1"].Formula = "T.INV.2T(0.45,0.99)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TInv2tIncorrectProbabilityOver()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM!");
                sheet.Cells["A1"].Formula = "T.INV.2T(1.06,3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TInv2tIncorrectProbabilityUnder()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return #NUM!");
                sheet.Cells["A1"].Formula = "T.INV.2T(-0.5,3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
        [TestMethod]
        public void TInv2tCorrectReuslt2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV.2T(0.99,18)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(0.0127087, result);
            }
        }
        [TestMethod]
        public void TInv2tCorrectReuslt3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV.2T(0.12,9.678)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(1.7175789, result);
            }
        }

        [TestMethod]
        public void TInv2tCorrectReuslt4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.INV.2T(0.86,34)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.177716773, result);
            }
        }
    }
}
