using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class TextFunctionsArrayTests
    {
        
        [TestMethod]
        public void TrimShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = " data ";
                sheet.Cells["A2"].Value = "data2 ";
                sheet.Cells["A3"].Value = " data3";
                sheet.Cells["B1:B3"].CreateArrayFormula("TRIM(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual("data", sheet.Cells["B1"].Value);
                Assert.AreEqual("data2", sheet.Cells["B2"].Value);
                Assert.AreEqual("data3", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void LowerShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "DATA";
                sheet.Cells["A2"].Value = "data2";
                sheet.Cells["A3"].Value = "daTa3";
                sheet.Cells["B1:B3"].CreateArrayFormula("LOWER(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual("data", sheet.Cells["B1"].Value);
                Assert.AreEqual("data2", sheet.Cells["B2"].Value);
                Assert.AreEqual("data3", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void UpperShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "data";
                sheet.Cells["A2"].Value = "data2";
                sheet.Cells["A3"].Value = "daTa3";
                sheet.Cells["B1:B3"].CreateArrayFormula("UPPER(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual("DATA", sheet.Cells["B1"].Value);
                Assert.AreEqual("DATA2", sheet.Cells["B2"].Value);
                Assert.AreEqual("DATA3", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void LeftShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "data";
                sheet.Cells["A2"].Value = " data2";
                sheet.Cells["A3"].Value = "daTa3";
                sheet.Cells["B1:B3"].CreateArrayFormula("LEFT(A1:A3, 3)");
                sheet.Calculate();
                Assert.AreEqual("dat", sheet.Cells["B1"].Value);
                Assert.AreEqual(" da", sheet.Cells["B2"].Value);
                Assert.AreEqual("daT", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void RightShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "data";
                sheet.Cells["A2"].Value = "data2";
                sheet.Cells["A3"].Value = "daTa3";
                sheet.Cells["B1:B3"].CreateArrayFormula("RIGHT(A1:A3, 3)");
                sheet.Calculate();
                Assert.AreEqual("ata", sheet.Cells["B1"].Value);
                Assert.AreEqual("ta2", sheet.Cells["B2"].Value);
                Assert.AreEqual("Ta3", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void MidShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "abc";
                sheet.Cells["A2"].Value = "def";
                sheet.Cells["A3"].Value = "ghi";
                sheet.Cells["B1:B3"].CreateArrayFormula("MID(A1:A3,2,1)");
                sheet.Calculate();
                Assert.AreEqual("b", sheet.Cells["B1"].Value);
                Assert.AreEqual("e", sheet.Cells["B2"].Value);
                Assert.AreEqual("h", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void UnicodeShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "c";
                sheet.Cells["A3"].Value = "e";
                sheet.Cells["B1:B3"].CreateArrayFormula("UNICODE(A1:A3)");
                sheet.Calculate();
                Assert.AreEqual(97, sheet.Cells["B1"].Value);
                Assert.AreEqual(99, sheet.Cells["B2"].Value);
                Assert.AreEqual(101, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void NumberValueShouldReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "1";
                sheet.Cells["B1"].Value = "2";
                sheet.Cells["C1"].Value = "3";
                sheet.Cells["A2:C2"].CreateArrayFormula("NUMBERVALUE(A1:C1)");
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A2"].Value);
                Assert.AreEqual(2d, sheet.Cells["B2"].Value);
                Assert.AreEqual(3d, sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void LenReturnHorizontalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "test";
                sheet.Cells["B1"].Value = "zest";
                sheet.Cells["C1"].Value = "testing";
                sheet.Cells["A2:C2"].CreateArrayFormula("LEN(A1:C1)");
                sheet.Calculate();
                Assert.AreEqual(4d, sheet.Cells["A2"].Value);
                Assert.AreEqual(4d, sheet.Cells["B2"].Value);
                Assert.AreEqual(7d, sheet.Cells["C2"].Value);
            }
        }
    }
}
