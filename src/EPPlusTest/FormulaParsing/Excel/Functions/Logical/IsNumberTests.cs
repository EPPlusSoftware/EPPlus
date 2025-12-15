using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class IsNumberTests : TestBase
    {
        [TestMethod]
        public void IsNumberShouldReturnFalseOnBoolean()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "ISNUMBER(TRUE)";
            sheet.Cells["A2"].Formula = "ISNUMBER(FALSE)";
            sheet.Calculate();
            Assert.IsFalse((bool)sheet.Cells["A1"].Value);
            Assert.IsFalse((bool)sheet.Cells["A2"].Value);
        }
        [TestMethod]
        public void IsNumberShouldReturnTrueOnNumbers()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "ISNUMBER(1)";
            sheet.Cells["A2"].Formula = "ISNUMBER(B2)";
            sheet.Cells["B2"].Value = 3;
            sheet.Calculate();
            Assert.IsTrue((bool)sheet.Cells["A1"].Value);
            Assert.IsTrue((bool)sheet.Cells["A2"].Value);
        }
        [TestMethod]
        public void IsNumberShouldReturnFalseOnStringNumbers()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "ISNUMBER(\"1\")";
            sheet.Cells["A2"].Formula = "ISNUMBER(B2)";
            sheet.Cells["B2"].Value = "3";
            sheet.Calculate();
            Assert.IsFalse((bool)sheet.Cells["A1"].Value);
            Assert.IsFalse((bool)sheet.Cells["A2"].Value);
        }
    }
}

