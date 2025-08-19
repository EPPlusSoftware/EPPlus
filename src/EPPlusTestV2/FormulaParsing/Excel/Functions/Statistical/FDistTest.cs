using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class FDistTest : TestBase
    {

        [TestMethod]
        public void FDistWithRegularInputPDF()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.DIST(2.73,9,5,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.087132132d, result);

            }
        }

        [TestMethod]
        public void FDistWithRegularInputCDF()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.DIST(2.73,9,5,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.859402994d, result);

            }
        }

        [TestMethod]
        public void FDistIncorrectDf()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error: ");
                sheet.Cells["A1"].Formula = "F.DIST(2.73,9,0.87,TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void FDistNegativeX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error: ");
                sheet.Cells["A1"].Formula = "F.DIST(-2.73,9,5,TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);

            }
        }

        [TestMethod]
        public void FDistTruncate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel: ");
                sheet.Cells["A1"].Formula = "F.DIST(2.73,9.89,5,TRUE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.859402994d, result);

            }
        }

    }
}
