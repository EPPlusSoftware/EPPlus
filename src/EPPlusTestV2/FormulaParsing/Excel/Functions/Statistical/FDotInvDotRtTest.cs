using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class FDotInvDotRtTest
    {
        [TestMethod]
        public void FDotInvDotRt()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV.RT(0.2,9,5)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 6);
                Assert.AreEqual(2.196277d, result);
            }
        }

        [TestMethod]
        public void FDotInvDotRtShouldReturnErrorWrongProbability()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV.RT(1.2,9,5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void FDotInvDotRtShouldReturnErrorWrongDF1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV.RT(0.2,0.7,5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void FDotInvDotRtShouldReturnErrorWrongDF2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV.RT(0.2,0.7,0.8)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void FDotInvDotRtShouldReturnErrorTooHighDF2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.INV.RT(0.2,0.7,10000000000)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
    }
}
