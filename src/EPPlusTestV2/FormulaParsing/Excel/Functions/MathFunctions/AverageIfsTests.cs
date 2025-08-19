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
    public class AverageIfsTests : TestBase
    {
        [TestMethod]
        public void AverageIfsShouldNotCountNumericStringsAsNumbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = "2";
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[1, 3].Value = 2;
                sheet.Cells[2, 3].Value = 1;
                sheet.Cells[3, 3].Value = "4";

                sheet.Cells[4, 1].Formula = "AVERAGEIFS(A1:A3,B1:B3,\">0\",C1:C3,\">1\")";
                sheet.Calculate();
                var val = sheet.Cells[4, 1].Value;
                Assert.AreEqual(3d, val);
            }
        }
        [TestMethod]
        public void ShouldHandleErrorInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[1, 2].Value = "#REF!";
                sheet.Cells[2, 2].Value = new ExcelErrorValue(eErrorType.Ref); 
                sheet.Cells[3, 2].Value = 3;

                sheet.Cells[4, 1].Formula = "AVERAGEIFS(A1:A3, B1:B3, #REF!)";
                sheet.Calculate();
                var val = sheet.Cells[4, 1].Value;
                Assert.AreEqual(4d, val);
            }
        }

        [TestMethod]
        public void AverageIfsShouldIgnoreErrorsInRangeIfInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["A2"].Value = "a";
                sheet.Cells["B2"].Value = ErrorValues.NAError;
                sheet.Cells["C2"].Value = "Test";

                sheet.Cells["A3"].Formula = "AVERAGEIFS(A1:C1,A2:C2,\"=#N/A\")";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["A3"].Value);

                sheet.Cells["A3"].Formula = "AVERAGEIFS(A1:C1,A2:C2,\"=a\")";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A3"].Value);
            }
        }


        [TestMethod]
        public void AverageIfsShouldCountMatchingQuotedFalseValue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123;
                sheet.Cells["B1"].Value = false;
                sheet.Cells[2, 1].Formula = "AverageIfs(A1,B1,\"FALSE\")";
                sheet.Calculate();
                var val = sheet.Cells[2, 1].Value;
                Assert.AreEqual(123d, val);
            }
        }
        [TestMethod]
        public void AverageIfsShouldHandleArraysInTheCriteriaRange_ColumnWise()
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
                sheet.Cells["C2"].Formula = "=AVERAGEIFS(N2:N11,K2:K11,A2:A11)";
                sheet.Cells["D2"].Formula = "=AVERAGEIFS(N2:N11,K2:K11,A2:A3)";
                sheet.Cells["E2"].Formula = "=AVERAGEIFS(N2:N11,K2:K11,A2:B3)";

                sheet.Calculate();

                Assert.AreEqual("C2:C11", sheet.Cells["C2"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["C2"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["C3"].Value, 0.000001);
                Assert.AreEqual(33.12, (double)sheet.Cells["C4"].Value, 0.000001);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["C5"].Value);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["C11"].Value);
                Assert.IsNull(sheet.Cells["C12"].Value);

                Assert.AreEqual("D2:D3", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["D2"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["D3"].Value, 0.000001);
                Assert.IsNull(sheet.Cells["D4"].Value);

                Assert.AreEqual("E2:F3", sheet.Cells["E2"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["E2"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["E3"].Value, 0.000001);
                Assert.IsNull(sheet.Cells["D4"].Value);
                Assert.AreEqual(29.4, (double)sheet.Cells["F2"].Value, 0.000001);
                Assert.AreEqual(7.2, sheet.Cells["F3"].Value);
                Assert.IsNull(sheet.Cells["F4"].Value);
            }
        }
        [TestMethod]
        public void AverageIfsShouldHandleArraysInTheCriteriaRange_RowWise()
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
                sheet.Cells["C5"].Formula = "AVERAGEIFS(N2:N11,K2:K11,A2:F2)";
                sheet.Cells["C6"].Formula = "AVERAGEIFS(N2:N11,K2:K11,A2:B2)";
                sheet.Cells["C7"].Formula = "AVERAGEIFS(N2:N11,K2:K11,A2:B3)";

                sheet.Calculate();

                Assert.AreEqual("C5:H5", sheet.Cells["C5"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["C5"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["D5"].Value, 0.000001);
                Assert.AreEqual(33.12, (double)sheet.Cells["E5"].Value, 0.000001);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["F5"].Value);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["G5"].Value);
                Assert.IsNull(sheet.Cells["I5"].Value);

                Assert.AreEqual("C6:D6", sheet.Cells["C6"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["C6"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["D6"].Value, 0.000001);
                Assert.IsNull(sheet.Cells["E6"].Value);

                Assert.AreEqual("C7:D8", sheet.Cells["C7"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["C7"].Value, 0.000001);
                Assert.AreEqual(29.4, (double)sheet.Cells["D7"].Value, 0.000001);
                Assert.IsNull(sheet.Cells["E7"].Value);
                Assert.AreEqual(29.4, (double)sheet.Cells["C8"].Value, 0.000001);
                Assert.AreEqual(7.2, sheet.Cells["D8"].Value);
                Assert.IsNull(sheet.Cells["E8"].Value);
            }
        }
        [TestMethod]
        public void AverageIfsShouldHandleArraysWithMultipleCriteria()
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

                sheet.Cells["C2"].Formula = "AVERAGEIFS(N2:N11,K2:K11,A2:A5,L2:L11,B2:B4)";
                sheet.Cells["D2"].Formula = "AVERAGEIFS(N2:N11,K2:K11,A2:A5,N2:N11,\">50\")";

                sheet.Calculate();

                Assert.AreEqual("C2:C5", sheet.Cells["C2"].FormulaRange.Address);
                Assert.AreEqual(90.2, (double)sheet.Cells["C2"].Value, 0.000001);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["C3"].Value);
                Assert.AreEqual(33.12D, sheet.Cells["C4"].Value);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["C5"].Value);
                Assert.IsNull(sheet.Cells["D6"].Value);

                Assert.AreEqual("D2:D5", sheet.Cells["D2"].FormulaRange.Address);
                Assert.AreEqual(129.2, (double)sheet.Cells["D2"].Value, 0.000001);
                Assert.AreEqual(72.7D, sheet.Cells["D3"].Value);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["D4"].Value);
                Assert.AreEqual(ErrorValues.Div0Error, sheet.Cells["D5"].Value);
                Assert.IsNull(sheet.Cells["D6"].Value);

                SaveWorkbook("AverageIfsMultiArray.xlsx", package);
            }
        }
    }
}