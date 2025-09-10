using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class ByColTests 
    {
        [TestMethod]
        public void ByColTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;
            
            sheet.Cells["D4"].Formula = "BYCOL(A1:C2, LAMBDA(array, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(4d, sheet.Cells["D4"].Value);
            Assert.AreEqual(5d, sheet.Cells["E4"].Value);
            Assert.AreEqual(6d, sheet.Cells["F4"].Value);
        }

        [TestMethod]
        public void ByCol_InMemoryRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D4"].Formula = "BYCOL(A1:C2 + 1, LAMBDA(array, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(5d, sheet.Cells["D4"].Value);
            Assert.AreEqual(6d, sheet.Cells["E4"].Value);
            Assert.AreEqual(7d, sheet.Cells["F4"].Value);
        }

        [TestMethod]
        public void ByCol_ShouldReturnValueErrorIfWrongNumberOfParams()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D4"].Formula = "BYCOL(A1:C2, LAMBDA(array, a, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D4"].Value);
        }

        [TestMethod]
        public void ByCol_Xleta_Attribute_Average()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 12;
            sheet.Cells["A2"].Value = 46;
            sheet.Cells["A3"].Value = 23;
            sheet.Cells["A4"].Value = 60;
            sheet.Cells["C4"].Formula = "BYCOL(A1:A4, _xleta.AVERAGE)";
            sheet.Calculate();
            Assert.AreEqual(35.25d, sheet.Cells["C4"].Value);
        }
    }
}
