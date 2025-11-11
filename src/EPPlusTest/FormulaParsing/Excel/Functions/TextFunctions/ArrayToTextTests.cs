using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using FakeItEasy.Configuration;

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
            Assert.AreEqual("1; 3; 2; 4", result);
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
            Assert.AreEqual("{1\\3;2\\4}", result);
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
            Assert.AreEqual("1; 3; #DIV/0!; 4", result);
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
            Assert.AreEqual("{\"Hello\"\\3;2\\4}", result);
        }

        [TestMethod]
        public void ArrayToText_StringConcise()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var dt = new DateTime(2025, 04, 11);
            var oaDate = dt.ToOADate();
            sheet.Cells["A1"].Value = oaDate;
            //sheet.Cells["A1"].Value = "2025-04-11";
            sheet.Cells["B1"].Value = 1416;
            sheet.Cells["A2"].Value = true;
            sheet.Cells["B2"].Value = "Katt";
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B3, 0)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            Assert.AreEqual($"{oaDate}; 1416; TRUE; Katt; ;", result);
        }

        [TestMethod]
        public void ArrayToText_StringStrict()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var dt = new DateTime(2025, 04, 11);
            sheet.Cells["A1"].Value = "2025-04-11";
            sheet.Cells["B1"].Value = 1416;
            sheet.Cells["A2"].Value = true;
            sheet.Cells["B2"].Value = "Katt";
            sheet.Cells["C1"].Formula = "ARRAYTOTEXT(A1:B3, 1)";
            sheet.Calculate();
            var result = sheet.Cells["C1"].Value;
            var oaDate = dt.ToOADate();
            Assert.AreEqual("{\"2025-04-11\"\\1416;TRUE\\\"Katt\";\\}", result);
        }
        [TestMethod]
        public void ArrayToText_StringConcise2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var dt = new DateTime(2025, 04, 11);
            sheet.Cells["A1"].Value = true;
            sheet.Cells["B1"].Value = "Katt";
            sheet.Cells["A3"].Value = 1416;
            sheet.Cells["B3"].Value = true;
            sheet.Cells["A4"].Formula = "SUM(100+400)";
            sheet.Cells["B4"].Value = "\"KATTEN\"";
            sheet.Cells["B5"].Formula = "ARRAYTOTEXT(A1:B4, 0)";
            sheet.Calculate();
            var result = sheet.Cells["B5"].Value;
            var oaDate = dt.ToOADate();
            Assert.AreEqual("TRUE; Katt; ; ; 1416; TRUE; 500; \"KATTEN\"", result);
        }

        [TestMethod]
        public void ArrayToText_StringStrict2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var dt = new DateTime(2025, 04, 11);
            sheet.Cells["A1"].Value = true;
            sheet.Cells["B1"].Value = "Katt";
            sheet.Cells["A3"].Value = 1416;
            sheet.Cells["B3"].Value = true;
            sheet.Cells["A4"].Formula = "SUM(100+400)";
            sheet.Cells["B4"].Value = "\"KATTEN\"";
            sheet.Cells["B5"].Formula = "ARRAYTOTEXT(A1:B4, 1)";
            sheet.Calculate();
            var result = sheet.Cells["B5"].Value;
            var oaDate = dt.ToOADate();
            Assert.AreEqual("{TRUE\\\"Katt\";\\;1416\\TRUE;500\\\"\"\"KATTEN\"\"\"}", result);
        }
    }
}
