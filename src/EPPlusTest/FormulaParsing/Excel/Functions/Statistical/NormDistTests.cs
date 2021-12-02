using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class NormDistTests
    {
        [TestMethod]
        public void NormDistShouldReturnCorrectResult()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Formula = "NORMDIST(1, 2, 3, TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.36944134d, System.Math.Round((double)result, 8));

                sheet.Cells["A2"].Formula = "NORMDIST(1, 2, 3, FALSE)";
                sheet.Calculate();
                result = sheet.Cells["A2"].Value;
                Assert.AreEqual(0.125794409d, System.Math.Round((double)result, 9));
            }
        }

        [TestMethod]
        public void NormDotDistShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Formula = "NORM.DIST(1.5, 2.345, 3, TRUE)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.389099558, System.Math.Round((double)result, 9));

                sheet.Cells["A2"].Formula = "NORM.DIST(1.5, 2.345, 3, FALSE)";
                sheet.Calculate();
                result = sheet.Cells["A2"].Value;
                Assert.AreEqual(0.127808928d, System.Math.Round((double)result, 9));
            }
        }
    }
}
