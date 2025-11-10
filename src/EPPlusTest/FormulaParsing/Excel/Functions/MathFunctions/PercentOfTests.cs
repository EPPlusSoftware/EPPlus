using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class PercentOfTests : TestBase
    {
        [TestMethod]
        public void PercentOfTestCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = 20.5;
                sheet.Cells["A2"].Value = 30.234;
                sheet.Cells["A3"].Value = 3.21312;                
                sheet.Cells["A4"].Value = 543;
                sheet.Cells["A5"].Value = 45.32;
                sheet.Cells["B1"].Formula = "PERCENTOF(A1:A2, A1:A5)";
                sheet.Calculate();                
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(0.0789920555173368, (double)result, 0.000000000001);
            }
        }

        [TestMethod]
        public void PercentOfErrorTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "cookie";
                sheet.Cells["A2"].Value = "snail";
                sheet.Cells["A3"].Value = 30.234;
                sheet.Cells["B1"].Formula = "PERCENTOF(A3, A1:A2)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result);
            }
        }

        [TestMethod]
        public void PercentOfTest3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["B1"].Formula = "PERCENTOF(1, 1)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(1d, result);
            }
        }

        [TestMethod]
        public void PercentOfShouldHandleText()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "cookie";
                sheet.Cells["A2"].Value = "snail";
                sheet.Cells["A3"].Value = 30.234;
                sheet.Cells["B1"].Formula = "PERCENTOF(A1:A2, A3)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(0d, result);
            }
        }

        [TestMethod]
        public void PercentOfShouldIgnoreNumericString()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "123";
                sheet.Cells["A2"].Value = "snail";
                sheet.Cells["A3"].Value = 30.234;
                sheet.Cells["B1"].Formula = "PERCENTOF(A1:A2, A3)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(0d, result);
            }
        }

        [TestMethod]
        public void PercentOfShouldReturnNum()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = ErrorValues.RefError;
                sheet.Cells["A2"].Value = "snail";
                sheet.Cells["A3"].Value = 30.234;
                sheet.Cells["B1"].Formula = "PERCENTOF(A1:A2, A3)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void PercentOfShouldReturnNum2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = 2.1;
                sheet.Cells["A2"].Value = ErrorValues.RefError;
                sheet.Cells["A3"].Value = 30.234;
                sheet.Cells["B1"].Formula = "PERCENTOF(A1, A2:A3)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }
    }
 }
