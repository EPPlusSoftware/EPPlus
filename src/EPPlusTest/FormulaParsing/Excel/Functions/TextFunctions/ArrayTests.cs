using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class ArrayTests
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
    }
}
