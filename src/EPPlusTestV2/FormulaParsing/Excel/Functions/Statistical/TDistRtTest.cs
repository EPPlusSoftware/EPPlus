using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TDistRtTest : TestBase
    {
        [TestMethod]
        public void TDistRtCorrectReuslt()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST.RT(4,8)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.001974886, result);
            }
        }

        [TestMethod]
        public void TDistRtNegativeX()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST.RT(-4.43,3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 9);
                Assert.AreEqual(0.989314562, result);
            }
        }

        [TestMethod]
        public void TDistDegreesOfFreedomLessThanOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error");
                sheet.Cells["A1"].Formula = "T.DIST.RT(-4.43,0.9)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TDistRtExample()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return same result as excel");
                sheet.Cells["A1"].Formula = "T.DIST.RT(16.64,3.3)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 13);
                Assert.AreEqual(0.0002362451567, result);
            }
        }
    }
}
