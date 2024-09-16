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
    public class CountIfsTests
    {
        [TestMethod]
        public void CountIfsShouldNotCountNumericStringsAsNumbers()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = "123";
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\">0\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingNumericValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 123;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\">0\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingQuotedFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = false;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\"FALSE\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingRawFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = false;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,FALSE)";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingQuotedTrueValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = true;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\"TRUE\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldCountMatchingRawTrueValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = true;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,TRUE)";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(1d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldNotCountMatchingQuotedZeroAsFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = false;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,\"0\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void CountIfsShouldNotCountMatchingRawZeroAsFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = false;
                sheet.Cells[2, 1].Formula = "COUNTIFS(A1,0)";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(0d, val);
            }
        }

        [TestMethod]
        public void CountIfsShoulCountRecordsMatchingAllCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells[1, 1].Value = 10;
                sheet.Cells[1, 2].Value = true;
                sheet.Cells[2, 1].Value = 15;
                sheet.Cells[2, 2].Value = true;
                sheet.Cells[3, 1].Value = 20;
                sheet.Cells[3, 2].Value = false;

                sheet.Cells[5, 1].Formula = "COUNTIFS(A1:A3,\"<20\",B1:B3,\"true\")";
                sheet.Cells[6, 1].Formula = "COUNTIFS(A1:A3,\">14\",B1:B3,\"true\")";
                sheet.Cells[7, 1].Formula = "COUNTIFS(A1:A3,\">=10\",B1:B3,\"false\")";
                sheet.Calculate();
                var val5 = sheet.Cells[5, 1].Value;
                Assert.AreEqual(2d, val5);
                var val6 = sheet.Cells[6, 1].Value;
                Assert.AreEqual(1d, val6);
                var val7 = sheet.Cells[7, 1].Value;
                Assert.AreEqual(1d, val7);
            }
        }
    }
}
