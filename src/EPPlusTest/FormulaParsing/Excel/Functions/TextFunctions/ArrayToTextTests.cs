using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class ArrayToTextTests
    {
        [TestMethod]
        public void ArrayToText_NumericConsiseFormat()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["B1"].Value = 3;
            sheet.Cells["B2"].Value = 4;
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B2, 0)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("1,2,3,4", result);
        }

        [TestMethod]
        public void ArrayToText_NumericStrictFormat()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["B1"].Value = 3;
            sheet.Cells["B2"].Value = 4;
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B2, 1)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("{1,2;3,4}", result);
        }

        [TestMethod]
        public void ArrayToText_ShouldIncludeError()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Formula = "1/0";
            sheet.Cells["B1"].Value = 3;
            sheet.Cells["B2"].Value = 4;
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B2, 0)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("1,#DIV/0!,3,4", result);
        }

        [TestMethod]
        public void ArrayToText_StringStrictFormat()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Hello";
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["B1"].Value = 3;
            sheet.Cells["B2"].Value = 4;
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B2, 1)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual("{\"Hello\",2;3,4}", result);
        }
    }
}
