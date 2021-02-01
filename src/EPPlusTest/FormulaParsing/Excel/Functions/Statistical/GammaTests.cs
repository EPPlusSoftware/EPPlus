using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class GammaTests
    {
        [TestMethod]
        public void GammalnShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "GAMMALN(4.5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(2.453736571, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "GAMMALN.PRECISE(4.5)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(2.453736571, System.Math.Round((double)result, 9));
            }
        }

        [TestMethod]
        public void GammaShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "GAMMA(1)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(1d, result);

                sheet.Cells["A1"].Formula = "GAMMA(5.5)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(52.34277778, System.Math.Round((double)result, 8));

            }
        }
    }
}
