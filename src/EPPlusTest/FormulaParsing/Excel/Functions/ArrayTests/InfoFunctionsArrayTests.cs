using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.ArrayTests
{
    [TestClass]
    public class InfoFunctionsArrayTests
    {
        [TestMethod]
        public void ErrorTypeShouldReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = ErrorValues.Div0Error;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["C1"].Value = ErrorValues.NameError;
                sheet.Cells["A2:C2"].CreateArrayFormula("ERROR.TYPE(A1:C1)");
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["B2"].Value);
                Assert.AreEqual(5, sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void IsBlankShouldReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = null;
                sheet.Cells["C1"].Value = 2;
                sheet.Cells["A2:C2"].CreateArrayFormula("ISBLANK(A1:C1)");
                sheet.Calculate();
                Assert.IsFalse((bool)sheet.Cells["A2"].Value);
                Assert.IsTrue((bool)sheet.Cells["B2"].Value);
                Assert.IsFalse((bool)sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void IsErrorShouldReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = ErrorValues.Div0Error;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["C1"].Value = ErrorValues.NameError;
                sheet.Cells["A2:C2"].CreateArrayFormula("ISERROR(A1:C1)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void IsErrShouldReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = ErrorValues.Div0Error;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["C1"].Value = ErrorValues.NameError;
                sheet.Cells["A2:C2"].CreateArrayFormula("ISERR(A1:C1)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void IsLogicalShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = true;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = false;
                sheet.Cells["B1:B3"].CreateArrayFormula("ISLOGICAL(A1:A3)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["B1"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsNaShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = ErrorValues.NAError;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["A3"].Value = ErrorValues.Div0Error;
                sheet.Cells["B1:B3"].CreateArrayFormula("ISNA(A1:A3)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["B1"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsFalse((bool)sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsNonTextShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = ErrorValues.NAError;
                sheet.Cells["A2"].Value = "Hello";
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["B1:B3"].CreateArrayFormula("ISNONTEXT(A1:A3)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["B1"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsNumberShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = ErrorValues.NAError;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1:B3"].CreateArrayFormula("ISNUMBER(A1:A3)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["B1"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsTextShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "abc";
                sheet.Cells["A2"].Value = ErrorValues.NAError;
                sheet.Cells["A3"].Value = "def";
                sheet.Cells["B1:B3"].CreateArrayFormula("ISTEXT(A1:A3)");
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["B1"].Value);
                Assert.IsFalse((bool)sheet.Cells["B2"].Value);
                Assert.IsTrue((bool)sheet.Cells["B3"].Value);
            }
        }
    }
}
