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
        public void ImAbsShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSUM(\"3+5i, i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                var roundedResult = System.Math.Round(System.Convert.ToDouble(result), 0);
                Assert.AreEqual("3+6i", roundedResult);
            }
        }
    }
}
