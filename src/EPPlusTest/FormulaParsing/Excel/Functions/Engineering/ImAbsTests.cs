using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ImAbsTests
    {
        [TestMethod]
        public void ImAbsShouldReturnCorrectResult() 
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMABS(\"5-2j\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                var roundedResult = System.Math.Round(System.Convert.ToDouble(result), 8);
                Assert.AreEqual(5.38516481, roundedResult);
            }
        }
        [TestMethod]
        public void ImAbsShouldReturnCorrectResultWithNumber()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMABS(14)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                var roundedResult = System.Math.Round(System.Convert.ToDouble(result), 6);
                Assert.AreEqual(14D, roundedResult);
            }
        }
        [TestMethod]
        public void ImAbsShouldReturnNumError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMABS(\"8\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                var roundedResult = System.Math.Round(System.Convert.ToDouble(result), 6);
                Assert.AreEqual(14.422205, roundedResult);
            }
        }

    }
}
