using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class LenTests
    {
        [TestMethod]
        public void LenShouldReturnCorrect()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "data";
                sheet.Cells["A2"].Formula = "LEN(A1)";
                sheet.Cells["A2"].Calculate();
                Assert.AreEqual(4d, sheet.Cells["A2"].Value, "LEN returned incorrect result when reading data from cell");

                sheet.Cells["A2"].Formula = "LEN(B1)";
                sheet.Cells["A2"].Calculate();
                Assert.AreEqual(0d, sheet.Cells["A2"].Value, "LEN returned incorrect result when reading null value from cell");

                sheet.Cells["A2"].Formula = "LEN(\"data\")";
                sheet.Cells["A2"].Calculate();
                Assert.AreEqual(4d, sheet.Cells["A2"].Value, "LEN returned incorrect result when reading hardcoded string");

            }
        }
    }
}
