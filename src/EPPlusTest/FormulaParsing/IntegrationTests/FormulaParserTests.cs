using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
        string QStr(string s)
        {
            char quotechar = '\"';
            return $"{quotechar}{s}{quotechar}";
        }
    }
}
