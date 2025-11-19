using OfficeOpenXml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class ValueToTextTests: TestBase
    {
        [TestMethod]
        public void ValueToTextTestSingleCell()
        {
            SwitchToCulture();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["C1"].Formula = "VALUETOTEXT(A1, 0)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("1", result);
            SwitchBackToCurrentCulture();
        }

        [TestMethod]
        public void ValueToTextSingleCellStrict()
        {
            SwitchToCulture();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "katt";
            sheet.Cells["C1"].Formula = "VALUETOTEXT(A1, 1)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("\"katt\"", result);
            SwitchBackToCurrentCulture();
        }

        [TestMethod]
        public void ValueToTextRangeConcise()
        {
            SwitchToCulture();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "katt";
            sheet.Cells["B1"].Value = true;
            sheet.Cells["A2"].Value = 123.1;
            sheet.Cells["B2"].Formula = "1/0";
            sheet.Cells["C1"].Formula = "VALUETOTEXT(A1:B2, 0)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            var result2 = sheet.Cells["D1"].Value;
            var result3 = sheet.Cells["C2"].Value;
            var result4 = sheet.Cells["D2"].Value;
            Assert.AreEqual("katt", result);
            Assert.AreEqual("TRUE", result2);
            Assert.AreEqual("123,1", result3);
            Assert.AreEqual("#DIV/0!", result4);
            SwitchBackToCurrentCulture();
        }

        [TestMethod]
        public void ValueToTextRangeStrict()
        {
            SwitchToCulture();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "katt";
            sheet.Cells["B1"].Value = true;
            sheet.Cells["A2"].Value = 123.1;
            sheet.Cells["B2"].Formula = "1/0";
            sheet.Cells["C1"].Formula = "VALUETOTEXT(A1:B2, 1)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            var result2 = sheet.Cells["D1"].Value;
            var result3 = sheet.Cells["C2"].Value;
            var result4 = sheet.Cells["D2"].Value;
            Assert.AreEqual("\"katt\"", result);
            Assert.AreEqual("TRUE", result2);
            Assert.AreEqual("123,1", result3);
            Assert.AreEqual("#DIV/0!", result4);
            SwitchBackToCurrentCulture();
        }
    }
}
