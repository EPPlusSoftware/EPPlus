using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class SubtotalTests : TestBase
    {
        [TestMethod]
        public void IgnoringNested1To10()
        {
            using(var package = OpenTemplatePackage("Subtotal_All.xlsx"))
            {
                var sheet1 = package.Workbook.Worksheets["Sheet1"];

                sheet1.Calculate();
                
                Assert.AreEqual(5.5, sheet1.Cells["A6"].Value, "AVERAGE (1)");
                Assert.AreEqual(5.5, sheet1.Cells["A8"].Value, "AVERAGE (1) nested");

                Assert.AreEqual(3d, sheet1.Cells["B6"].Value, "COUNT (2)");
                Assert.AreEqual(3d, sheet1.Cells["B8"].Value, "COUNT (2) nested");

                Assert.AreEqual(3d, sheet1.Cells["C6"].Value, "COUNTA (3)");
                Assert.AreEqual(3d, sheet1.Cells["C8"].Value, "COUNTA (3) nested");

                Assert.AreEqual(8d, sheet1.Cells["D6"].Value, "MAX (4)");
                Assert.AreEqual(8d, sheet1.Cells["D8"].Value, "MAX (4) nested");

                Assert.AreEqual(2d, sheet1.Cells["E6"].Value, "MIN (5)");
                Assert.AreEqual(2d, sheet1.Cells["E8"].Value, "MIN (5) nested");

                Assert.AreEqual(280d, sheet1.Cells["F6"].Value, "PRODUCT (6)");
                Assert.AreEqual(280d, sheet1.Cells["F8"].Value, "PRODUCT (6) nested");

                EppAssert.DoublesAreEqual(2.9439, sheet1.Cells["G6"].Value, 4, "STDEV.S (7)");
                EppAssert.DoublesAreEqual(2.9439, sheet1.Cells["G8"].Value, 4, "STDEV.S (7) nested");

                EppAssert.DoublesAreEqual(1.118, sheet1.Cells["H6"].Value, 4, "STDEV.P (8)");
                EppAssert.DoublesAreEqual(1.118, sheet1.Cells["H8"].Value, 4, "STDEV.P (8) nested");

                Assert.AreEqual(10d, sheet1.Cells["I6"].Value, "SUM (9)");
                Assert.AreEqual(10d, sheet1.Cells["I8"].Value, "SUM (9) nested");

                EppAssert.DoublesAreEqual(6.9167, sheet1.Cells["J6"].Value, 4, "VAR.S (10)");
                EppAssert.DoublesAreEqual(6.9167, sheet1.Cells["J8"].Value, 4, "VAR.S (10) nested");

                EppAssert.DoublesAreEqual(2.6875, sheet1.Cells["K6"].Value, 4, "VAR.P (11)");
                EppAssert.DoublesAreEqual(2.6875, sheet1.Cells["K8"].Value, 4, "VAR.P (11) nested");

                Assert.AreEqual("OK", sheet1.Cells["D11"].Value, "Test with IF and nested SUBTOTAL failed");
            }
        }

        [TestMethod]
        public void ThreeNestedSubtotals()
        {
            using (var package = OpenTemplatePackage("Subtotal_All.xlsx"))
            {
                package.Workbook.Calculate();
                var sheet1 = package.Workbook.Worksheets["Sheet1"];

                Assert.AreEqual(22d, sheet1.Cells["J11"].Value);
            }
        }

        [TestMethod]
        public void SumTest()
        {
            using (var package = OpenTemplatePackage("Subtotal_All.xlsx"))
            {
                package.Workbook.Calculate();
                var sheet1 = package.Workbook.Worksheets["Sheet1"];

                Assert.AreEqual(581d, sheet1.Cells["G11"].Value);
            }
        }
    }
}
