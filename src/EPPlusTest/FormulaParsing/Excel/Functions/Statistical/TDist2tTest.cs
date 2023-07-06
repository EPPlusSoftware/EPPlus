using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TDist2tTest : TestBase
    {
        [TestMethod]
        public void TDist2tShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST.2T(4.6798,15.321)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(0.0002963, result);
            }
        }

        [TestMethod]
        public void TDist2TNegativeX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error, since this is not allowed");
                sheet.Cells["A1"].Formula = "T.DIST.2T(-4.6798,15.321)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TDist2TDfLessThanOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error, since this is not allowed");
                sheet.Cells["A1"].Formula = "T.DIST.2T(4.6798,0.5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TDist2TNonNumeric()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error, since this is not allowed");
                sheet.Cells["A1"].Formula = "T.DIST.2T(\"hgjkdw\",0.5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void TDist2TDocumentationTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST.2T(1.959999998,60)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.05464493, result);
            }
        }

    }
}
