using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class DevsqTests
    {
        [TestMethod]
        public void DevsqShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = 5;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 9;
                sheet.Cells["A6"].Value = 7;

                sheet.Cells["B1"].Formula = "DEVSQ(A1:A6)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(47.5, result);

                sheet.Cells["B1"].Formula = "DEVSQ(A1:A6, 8, 10)";
                sheet.Calculate();
                result = sheet.Cells["B1"].Value;
                Assert.AreEqual(79.875, result);

            }
        }
    }
}
