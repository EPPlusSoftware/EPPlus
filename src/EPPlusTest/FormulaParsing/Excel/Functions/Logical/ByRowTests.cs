using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class ByRowTests
    {
        [TestMethod]
        public void ByRowTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D4"].Formula = "BYROW(A1:C2, LAMBDA(array, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(3d, sheet.Cells["D4"].Value);
            Assert.AreEqual(6d, sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void ByRow_InMemoryRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D4"].Formula = "BYROW(A1:C2 + 1, LAMBDA(array, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(4d, sheet.Cells["D4"].Value);
            Assert.AreEqual(7d, sheet.Cells["D5"].Value);
        }

        [TestMethod]
        public void ByRow_ShouldReturnValueErrorIfWrongNumberOfParams()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            sheet.Cells["A1"].Value = 1;
            sheet.Cells["B1"].Value = 2;
            sheet.Cells["C1"].Value = 3;
            sheet.Cells["A2"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["C2"].Value = 6;

            sheet.Cells["D4"].Formula = "BYROW(A1:C2, LAMBDA(array, a, MAX(array)))";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D4"].Value);
        }
    }
}
