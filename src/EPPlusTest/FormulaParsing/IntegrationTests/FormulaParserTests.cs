using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class FormulaParserTests : FormulaParserTestBase
    {
        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void ValidateFormulaParserWithIsWorksheets1Based(bool isWorksheets1Based)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                package.Compatibility.IsWorksheets1Based = isWorksheets1Based;

                var wb = package.Workbook;
                var sheetA = wb.Worksheets.Add("A");
                var sheetB = wb.Worksheets.Add("B");


                sheetA.SetValue("A1", 1);
                sheetA.SetFormula(1, 2, "A1 ");
                sheetA.SetFormula(1, 3, "B!A1");
                sheetA.Names.AddFormula("sheetANameToA", "A!A1");
                sheetA.Names.AddFormula("sheetANameToB", "B!A1");

                sheetB.SetValue("A1", 2);
                sheetB.SetFormula(1, 2, "A1");
                sheetB.SetFormula(1, 3, "A!A1");
                sheetB.Names.AddFormula("sheetBNameToA", "A!A1");
                sheetB.Names.AddFormula("sheetBNameToB", "B!A1");

                wb.Calculate();

                Assert.AreEqual(1, sheetA.GetValue(1, 1));
                Assert.AreEqual(1, sheetA.GetValue(1, 2));
                Assert.AreEqual(2, sheetA.GetValue(1, 3));
                Assert.AreEqual(1, sheetA.Names[0].Value);
                Assert.AreEqual(2, sheetA.Names[1].Value);

                Assert.AreEqual(2, sheetB.GetValue(1, 1));
                Assert.AreEqual(2, sheetB.GetValue(1, 2));
                Assert.AreEqual(1, sheetB.GetValue(1, 3));
                Assert.AreEqual(1, sheetB.Names[0].Value);
                Assert.AreEqual(2, sheetB.Names[1].Value);
            }
        }
        [TestMethod]
        public void CalculateSingleFormulaShouldNotSetCellValue()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.SetValue(1, 1, 1);
                ws.SetFormula(2, 1, "A1+1");

                var v=package.Workbook.FormulaParser.Parse("A1+A2", "Sheet1!A3", new ExcelCalculationOption() { FollowDependencyChain=false });
                
                Assert.AreEqual(1D, v);
                Assert.IsNull(ws.Cells["A3"].Value);

                v = package.Workbook.FormulaParser.Parse("A1+A2", "Sheet1!A3");
                Assert.AreEqual(3D, v);
                Assert.IsNull(ws.Cells["A3"].Value);
            }
        }
        [TestMethod]
        public void CalculateMultipleNegationTests()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.SetValue(1, 1, 1);
                ws.SetFormula(2, 1, "(-(-(---A1)))");
                ws.Calculate();

                Assert.AreEqual(-1D, ws.GetValue(2,1));
            }
        }
        string QStr(string s)
        {
            char quotechar = '\"';
            return $"{quotechar}{s}{quotechar}";
        }
        [TestMethod]
        public void ReferencingWorksheetThatDoesNotExistShouldReturnRef()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet ws = p.Workbook.Worksheets.Add("sheet1");

                ws.Cells["A1"].Formula = "Sheet2!A1";

                ws.Calculate();

                var value = ws.Cells["A1"].Value;
                Assert.AreEqual(ErrorValues.RefError, value);
            }
        }
    }
}
