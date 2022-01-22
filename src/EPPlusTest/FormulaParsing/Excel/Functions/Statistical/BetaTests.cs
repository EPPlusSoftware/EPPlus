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
    public class BetaTests
    {
        [TestMethod]
        public void BetaDotInvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Formula = "BETA.INV(0.685470581, 8, 10, 1, 3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(2d, System.Math.Round((double)result, 8));

                sheet.Cells["A1"].Formula = "BETA.INV(0.55,3,4)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.445812d, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void BetainvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Formula = "BETAINV(0.685470581, 8, 10, 1, 3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(2d, System.Math.Round((double)result, 8));

                sheet.Cells["A1"].Formula = "BETAINV(0.55,3,4)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.445812d, System.Math.Round((double)result, 6));
            }
        }
    }
}
