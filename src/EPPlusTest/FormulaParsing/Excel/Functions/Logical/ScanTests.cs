using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class ScanTests : TestBase
    {
        [TestMethod]
        public void ScanTestFactorial()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D5"].Formula = "SCAN(1,A1:C2,LAMBDA(a,b,a*b))";

            sheet.Calculate();

            Assert.AreEqual(1d, sheet.Cells["D5"].Value);
            Assert.AreEqual(2d, sheet.Cells["E5"].Value);
            Assert.AreEqual(6d, sheet.Cells["F5"].Value);
            Assert.AreEqual(24d, sheet.Cells["D6"].Value);
            Assert.AreEqual(120d, sheet.Cells["E6"].Value);
            Assert.AreEqual(720d, sheet.Cells["F6"].Value);
        }

        [TestMethod]
        public void ScanTestStringValues()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "a";
            sheet.Cells["B1"].Value = "b";
            sheet.Cells["C1"].Value = "c";
            sheet.Cells["A2"].Value = "d";
            sheet.Cells["B2"].Value = "e";
            sheet.Cells["C2"].Value = "f";

            sheet.Cells["D5"].Formula = "SCAN(\"\",A1:C2,LAMBDA(a,b,a&b))";

            sheet.Calculate();

            Assert.AreEqual("a", sheet.Cells["D5"].Value);
            Assert.AreEqual("ab", sheet.Cells["E5"].Value);
            Assert.AreEqual("abc", sheet.Cells["F5"].Value);
            Assert.AreEqual("abcd", sheet.Cells["D6"].Value);
            Assert.AreEqual("abcde", sheet.Cells["E6"].Value);
            Assert.AreEqual("abcdef", sheet.Cells["F6"].Value);
        }

        [TestMethod]
        public void ScanTestStringValues_OmitFirstArg()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "a";
            sheet.Cells["B1"].Value = "b";
            sheet.Cells["C1"].Value = "c";
            sheet.Cells["A2"].Value = "d";
            sheet.Cells["B2"].Value = "e";
            sheet.Cells["C2"].Value = "f";

            sheet.Cells["D5"].Formula = "SCAN(,A1:C2,LAMBDA(a,b,a&b))";

            sheet.Calculate();

            Assert.AreEqual("a", sheet.Cells["D5"].Value);
            Assert.AreEqual("ab", sheet.Cells["E5"].Value);
            Assert.AreEqual("abc", sheet.Cells["F5"].Value);
            Assert.AreEqual("abcd", sheet.Cells["D6"].Value);
            Assert.AreEqual("abcde", sheet.Cells["E6"].Value);
            Assert.AreEqual("abcdef", sheet.Cells["F6"].Value);
        }

        [TestMethod]
        public void ScanTestNumericValues_OmitFirstArg()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D5"].Formula = "SCAN(,A1:C2,LAMBDA(a,b,a + b))";

            sheet.Calculate();

            Assert.AreEqual(1d, sheet.Cells["D5"].Value);
            Assert.AreEqual(3d, sheet.Cells["E5"].Value);
            Assert.AreEqual(6d, sheet.Cells["F5"].Value);
            Assert.AreEqual(10d, sheet.Cells["D6"].Value);
            Assert.AreEqual(15d, sheet.Cells["E6"].Value);
            Assert.AreEqual(21d, sheet.Cells["F6"].Value);
        }

        [TestMethod]
        public void Scan_ShouldReturnCalcErrorWhenFirstArgIsRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D5"].Formula = "SCAN(H1:I1,A1:C2,LAMBDA(a,b,a + b))";

            sheet.Calculate();

            Assert.IsInstanceOfType(sheet.Cells["D5"].Value, typeof(ExcelErrorValue));
            Assert.AreEqual(DataType.ExcelError, ((ExcelErrorValue)sheet.Cells["D5"].Value).AsCompileResult.DataType);
            var str = sheet.Cells["D5"].Value.ToString();
            Assert.AreEqual("#CALC!", str);

            //SaveWorkbook("ScanErrors.xlsx", package);
        }

        [TestMethod]
        public void Scan_Xleta_Attribute_Sum()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 5;
            sheet.Cells["A2"].Value = 10;
            sheet.Cells["A3"].Value = 15;

            sheet.Cells["C1"].Formula = "SCAN(0, A1:A3, _xleta.SUM)";
            sheet.Calculate();

            Assert.AreEqual(5d, sheet.Cells["C1"].Value);     // första värdet
            Assert.AreEqual(15d, sheet.Cells["C2"].Value);    // andra värdet
            Assert.AreEqual(30d, sheet.Cells["C3"].Value);    // tredje värdet
        }
    }
}
