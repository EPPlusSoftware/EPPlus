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
    public class SumIfsTests
    {
        [TestMethod]
        public void SumIfsShouldHandleSingleRange()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SUMIFS(H5,H5,\">0\",K5,\"> 0\")";
                sheet.Cells["H5"].Value = 1;
                sheet.Cells["K5"].Value = 1;
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SumIfsShouldNotCountNumericStringsAsNumbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = "123";
                sheet.Cells[2, 1].Formula = "SUMIFS(A1,B1,\">0\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void SumIfsShouldCountMatchingQuotedFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["B1"].Value = false;
                sheet.Cells[2, 1].Formula = "SUMIFS(A1,B1,\"FALSE\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(123d, val);
            }
        }

        [TestMethod]
        public void SumIfsShouldCountMatchingRawFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["B1"].Value = false;
                sheet.Cells[2, 1].Formula = "SUMIFS(A1,B1,FALSE)";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(123d, val);
            }
        }

        [TestMethod]
        public void SumIfsShouldNotCountMatchingQuotedZeroAsFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["B1"].Value = false;
                sheet.Cells[2, 1].Formula = "SUMIFS(A1,B1,\"0\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void SumIfsShouldNotCountMatchingRawZeroAsFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["B1"].Value = false;
                sheet.Cells[2, 1].Formula = "SUMIFS(A1,B1, 0)";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }
    }
}
