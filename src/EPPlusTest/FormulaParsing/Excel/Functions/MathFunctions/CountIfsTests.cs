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
    [DoNotParallelize]
    public class CountIfsTests : TestBase
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
        public void CountIfsShouldCountRecordsMatchingAllCriteria()
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
        [TestMethod]
        public void CountIfsShouldHandleArraysInTheCriteriaRange_ColumnWise()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["A4"].Value = "Saw";
                sheet.Cells["B2"].Value = "Hammer";
                sheet.Cells["B3"].Value = "Butter";
                sheet.Cells["C2"].Formula = "=COUNTIFS(K2:K11,A2:A11)";
                sheet.Cells["D2"].Formula = "=COUNTIFS(K2:K11,A2:A3)";
                sheet.Cells["E2"].Formula = "=COUNTIFS(K2:K11,A2:B3)";

                sheet.Calculate();
                SaveWorkbook("CountIfsMultiArray.xlsx", package);
                Assert.AreEqual("C2:C11", sheet.Cells["C2"].FormulaRange.Address);
                Assert.AreEqual(3D, sheet.Cells["C2"].Value);
                Assert.AreEqual(3D, sheet.Cells["C3"].Value);
                Assert.AreEqual(1D, sheet.Cells["C4"].Value);
                Assert.AreEqual(0D, sheet.Cells["C5"].Value);
                Assert.AreEqual(0D, sheet.Cells["C11"].Value);
                Assert.IsNull(sheet.Cells["C12"].Value);

                Assert.AreEqual("D2:D3", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(3D, sheet.Cells["D2"].Value);
                Assert.AreEqual(3D, sheet.Cells["D3"].Value);
                Assert.IsNull(sheet.Cells["D4"].Value);

                Assert.AreEqual("E2:F3", sheet.Cells["E2"].FormulaRange.Address);
                Assert.AreEqual(3D, sheet.Cells["E2"].Value);
                Assert.AreEqual(3D, sheet.Cells["E3"].Value);
                Assert.IsNull(sheet.Cells["D4"].Value);
                Assert.AreEqual(3D, sheet.Cells["F2"].Value);
                Assert.AreEqual(1D, sheet.Cells["F3"].Value);
                Assert.IsNull(sheet.Cells["F4"].Value);
            }
        }
        [TestMethod]
        public void CountIfsShouldHandleArraysInTheCriteriaRange_RowWise()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["B2"].Value = "Hammer";
                sheet.Cells["C2"].Value = "Saw";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["B3"].Value = "Butter";
                sheet.Cells["C5"].Formula = "COUNTIFS(K2:K11,A2:F2)";
                sheet.Cells["C6"].Formula = "COUNTIFS(K2:K11,A2:B2)";
                sheet.Cells["C7"].Formula = "COUNTIFS(K2:K11,A2:B3)";

                sheet.Calculate();

                Assert.AreEqual("C5:H5", sheet.Cells["C5"].FormulaRange.Address);
                Assert.AreEqual(3D, sheet.Cells["C5"].Value);
                Assert.AreEqual(3D, sheet.Cells["D5"].Value);
                Assert.AreEqual(1D, sheet.Cells["E5"].Value);
                Assert.AreEqual(0D, sheet.Cells["F5"].Value);
                Assert.AreEqual(0D, sheet.Cells["G5"].Value);
                Assert.IsNull(sheet.Cells["I5"].Value);

                Assert.AreEqual("C6:D6", sheet.Cells["C6"].FormulaRange.Address);
                Assert.AreEqual(3D, (double)sheet.Cells["C6"].Value, 0.000001);
                Assert.AreEqual(3D, (double)sheet.Cells["D6"].Value, 0.000001);
                Assert.IsNull(sheet.Cells["E6"].Value);

                Assert.AreEqual("C7:D8", sheet.Cells["C7"].FormulaRange.Address);
                Assert.AreEqual(3D, sheet.Cells["C7"].Value);
                Assert.AreEqual(3D, sheet.Cells["D7"].Value);
                Assert.IsNull(sheet.Cells["E7"].Value);
                Assert.AreEqual(3D, sheet.Cells["C8"].Value);
                Assert.AreEqual(1D, sheet.Cells["D8"].Value);
                Assert.IsNull(sheet.Cells["E8"].Value);
            }
        }
        [TestMethod]
        [DoNotParallelize]
        public void CountIfsShouldHandleArraysWithMultipleCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                LoadItemData(sheet);
                sheet.Cells["A2"].Value = "Crowbar";
                sheet.Cells["A3"].Value = "Hammer";
                sheet.Cells["A4"].Value = "Saw";
                sheet.Cells["A5"].Value = "Monkey Wrench";
                sheet.Cells["B2"].Value = "Hardware";
                sheet.Cells["B3"].Value = "Software";
                sheet.Cells["B4"].Value = "Hardware";

                sheet.Cells["C2"].Formula = "COUNTIFS(K2:K11,A2:A5,L2:L11,B2:B4)";
                sheet.Cells["D2"].Formula = "COUNTIFS(K2:K11,A2:A5,N2:N11,\">50\")";

                sheet.Calculate();

                Assert.AreEqual("C2:C5", sheet.Cells["C2"].FormulaRange.Address);
                Assert.AreEqual(3D, (double)sheet.Cells["C2"].Value);
                Assert.AreEqual(0D, sheet.Cells["C3"].Value);
                Assert.AreEqual(1D, sheet.Cells["C4"].Value);
                Assert.AreEqual(0D, sheet.Cells["C5"].Value);
                Assert.IsNull(sheet.Cells["D6"].Value);

                Assert.AreEqual("D2:D5", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(2D, (double)sheet.Cells["D2"].Value, 0.000001);
                Assert.AreEqual(1D, sheet.Cells["D3"].Value);
                Assert.AreEqual(0D, sheet.Cells["D4"].Value);
                Assert.AreEqual(0D, sheet.Cells["D5"].Value);
                Assert.IsNull(sheet.Cells["D6"].Value);

                SaveWorkbook("CountIfsMultiArray.xlsx", package);
            }
        }
    }
}
