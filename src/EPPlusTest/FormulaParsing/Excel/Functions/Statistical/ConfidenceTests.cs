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
    public class ConfidenceTests
    {
        [TestMethod]
        public void ConfidenceNormShouldReturnCorrectResult()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "CONFIDENCE.NORM(0.05,0.36,100)";
                sheet.Cells["A2"].Formula = "CONFIDENCE.NORM(0.07,0.36,100)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 6);
                Assert.AreEqual(result, 0.070559d);
                result = System.Math.Round((double)sheet.Cells["A2"].Value, 6);
                Assert.AreEqual(result, 0.065229d);
            }
        }

        [TestMethod]
        public void ConfidenceTShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "CONFIDENCE.T(0.05,0.36,100)";
                sheet.Cells["A2"].Formula = "CONFIDENCE.T(0.07,0.36,100)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 6);
                Assert.AreEqual(result, 0.071432d);
                result = System.Math.Round((double)sheet.Cells["A2"].Value, 6);
                Assert.AreEqual(result, 0.065942d);
            }
        }
    }
}
