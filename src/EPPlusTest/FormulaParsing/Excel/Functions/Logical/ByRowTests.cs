using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class ByRowTests : TestBase
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

        [TestMethod]
        public void ByRow_WithLetAndLambda()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Formula = "BYROW(A1:A3, LAMBDA(row, LET(x, row, x + 1)))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["B1"].Value);
            Assert.AreEqual(3d, sheet.Cells["B2"].Value);
            Assert.AreEqual(4d, sheet.Cells["B3"].Value);
        }

        [TestMethod]
        public void ByRow_WithLetAndLambda2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Formula = "BYROW(A1:A3, LAMBDA(row, SUM(row)))";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["B1"].Value);
            Assert.AreEqual(2d, sheet.Cells["B2"].Value);
            Assert.AreEqual(3d, sheet.Cells["B3"].Value);
        }

        [TestMethod]
        public void ByRow_RangeRow()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            sheet.Cells["A3"].Value = 3;
            sheet.Cells["B1"].Value = 4;
            sheet.Cells["B2"].Value = 5;
            sheet.Cells["B3"].Value = 6;
            sheet.Cells["C1"].Formula = "BYROW(A1:B3, LAMBDA(row, SUM(row)))";
            sheet.Calculate();
            Assert.AreEqual(5d, sheet.Cells["C1"].Value);
            Assert.AreEqual(7d, sheet.Cells["C2"].Value);
            Assert.AreEqual(9d, sheet.Cells["C3"].Value);
        }

        [TestMethod]
        public void ByRow_Xleta_Attribute_Sum()
        { 
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 10;
            sheet.Cells["A2"].Value = 20;
            sheet.Cells["B1"].Value = 11;
            sheet.Cells["B2"].Value = 21;
            sheet.Cells["C4"].Formula = "BYROW(A1:B2, _xleta.SUM)";
            sheet.Calculate();
            Assert.AreEqual(21d, sheet.Cells["C4"].Value);
            Assert.AreEqual(41d, sheet.Cells["C5"].Value);
        }

        [TestMethod]
        public void ByRow_Xleta_Attribute_CountA()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 10;
            sheet.Cells["A2"].Value = 20;
            sheet.Cells["B1"].Value = 11;
            sheet.Cells["B3"].Value = 21;
            sheet.Cells["C4"].Formula = "BYROW(A1:B3, _xleta.COUNTA)";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["C4"].Value);
            Assert.AreEqual(1d, sheet.Cells["C5"].Value);
            Assert.AreEqual(1d, sheet.Cells["C6"].Value);
        }

        [TestMethod]
        public void ByRow_Xleta_Attribute_Max()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = 5;
            sheet.Cells["A2"].Value = 8;
            sheet.Cells["B1"].Value = 12;
            sheet.Cells["B2"].Value = 3;
            sheet.Cells["C4"].Formula = "BYROW(A1:B2, _xleta.MAX)";
            sheet.Calculate();
            Assert.AreEqual(12d, sheet.Cells["C4"].Value);
            Assert.AreEqual(8d, sheet.Cells["C5"].Value);
        }
    }
}
