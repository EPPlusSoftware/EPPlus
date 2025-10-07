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
        [TestMethod]
        public void SumIfs_SumOutsideCriteriaShouldNotThrowCircularReferences()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = "SumResult";
                // This shouldn't be a circular reference, because the 1:1="SUMABLE" condition should filter out A2 before the 2:2 filter is applied
                sheet1.Cells["A2"].Formula = "SumIfs(3:3, 1:1,\"SUMABLE\",2:2,\"<>\")";

                sheet1.Cells["B2"].Value = 1;
                sheet1.Cells["C2"].Value = 2;
                sheet1.Cells["E2"].Value = 4;
                sheet1.Cells["F2"].Value = 5;
                sheet1.Cells["G2"].Value = 6;

                sheet1.Cells["C1"].Value = "SUMABLE";
                sheet1.Cells["D1"].Value = "SUMABLE";
                sheet1.Cells["E1"].Value = "SUMABLE";
                sheet1.Cells["G1"].Value = "SUMABLE";

                sheet1.Cells["B3"].Value = 1;
                sheet1.Cells["C3"].Value = 2;
                sheet1.Cells["E3"].Value = 4;
                sheet1.Cells["F3"].Value = 5;
                sheet1.Cells["G3"].Value = 6;

                pck.Workbook.Calculate();

                Assert.AreEqual(12D, sheet1.Cells["A2"].GetValue<double>(), 0.00);
            }
        }
    }
}
