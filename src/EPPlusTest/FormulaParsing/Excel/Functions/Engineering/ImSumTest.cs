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
    public class ImSumTest
    {
        [TestMethod]
        public void ImSumShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSUM(\"3+5i\", \"2+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("5+9i", result);
            }
        }

        [TestMethod]
        public void ImSumShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSUM(\"3+5i\", \"2+4i\", \"5+7i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("10+16i", result);
            }
        }

        [TestMethod]
        public void ImSumShouldReturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSUM(\"5+6i\", \"i\", \"3\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("8+7i", result);
            }
        }

        [TestMethod]
        public void ImSumShouldReturnCorrectResult4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSUM(\"5+6i\", \"4+8j\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }
    }
}
