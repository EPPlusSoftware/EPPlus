using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class ModeDotMultTest
    {
        [TestMethod]
        public void ModeMultShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 2;
                sheet.Cells["A7"].Value = 1;
                sheet.Cells["A8"].Value = 2;
                sheet.Cells["A9"].Value = 3;
                sheet.Cells["A10"].Value = 5;
                sheet.Cells["A11"].Value = 6;
                sheet.Cells["A12"].Value = 1;
                sheet.Cells["A14"].Formula = "MODE.MULT(A1:A12)";
                sheet.Calculate();
                Assert.AreEqual(3d, sheet.Cells["A14"].Value);
            }
        }
    }
}
