using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{

    [TestClass]
    public class CountaTests
    {
        private void SetValues(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "Value 1";
            worksheet.Cells["A2"].Value = "Value 2";
            worksheet.Cells["A3"].Value = "Value 3";
            worksheet.Cells["A6"].Value = "Test 2";
            worksheet.Cells["A7"].Value = "Value 1";
            worksheet.Cells["A8"].Value = "Value 2";
            worksheet.Cells["A9"].Value = "Value 3";
            worksheet.Cells["B4"].Formula = "COUNTA(B1:B3)";
            worksheet.Cells["B10"].Formula = "COUNTA(C7:C9)";
            worksheet.Cells["C7"].Formula = "IF(B7;\"\";A7)";
            worksheet.Cells["C8"].Formula = "IF(B8;\"\";A8)";
            worksheet.Cells["C9"].Formula = "IF(B9;\"\";A9)";
        }

        [DataTestMethod]
        [DataRow(null, null, null, 0d)]
        [DataRow("a", null, null, 1)]
        [DataRow(null, "b", null, 1)]
        [DataRow(null, "b", "c", 2)]
        [DataRow("a", "b", "c", 3)]
        [DataRow("", "", "", 3)]
        public void CountA_WithGivenCellValues_ShouldReturnExpectedCount(string value1, string value2, string value3, double expectedValue)
        {
            using(var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("CountA");
                SetValues(worksheet);

                worksheet.Cells["B1"].Value = value1;
                worksheet.Cells["B2"].Value = value2;
                worksheet.Cells["B3"].Value = value3;
                package.Workbook.Calculate();
                Assert.AreEqual(expectedValue, worksheet.Cells["B4"].Value);
            }
        }

        [TestMethod]
        public void CountA_WithFormulaError_ExpectErrorIncludedInCount()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("CountA");
                SetValues(worksheet);
                worksheet.Cells["B7"].Value = "test";
                worksheet.Cells["B8"].Value = false;
                worksheet.Cells["B9"].Value = false;
                package.Workbook.Calculate();
                Assert.AreEqual(3d, worksheet.Cells["B10"].Value);
            }
        }
    }
}
