using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class CriteriaRangeFunctionsTests : TestBase
    {
        [TestMethod]
        public void SumAndAverageIfFunctionsShouldNotReturnCircularReferenceIfCriteriaIsNoMet()
        {
            using var p = OpenPackage("i1687-2.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Fruit";
            ws.Cells["B1"].Value = "Employee";
            ws.Cells["A2:A3"].Value = "Apples";
            ws.Cells["A4:A5"].Value = "Artichokes";
            ws.Cells["A6:A7"].Value = "Bananas";
            ws.Cells["A8:A9"].Value = "Carrots";
            ws.Cells["B2,B4,B6,B8"].Value = "Mats";
            ws.Cells["B3,B5"].Value = "Jan";
            ws.Cells["B7,B9"].Value = "Ossian";

            ws.Cells["C2"].Value = 10D;
            ws.Cells["C3"].Value = 11D;
            ws.Cells["C4"].Formula = "AVERAGEIFS(C2:C9,A2:A9,\"=B*\",B2:B9,\"Ossian\")";
            ws.Cells["C5"].Formula = "SUMIFS(C2:C9,A2:A9,\"=A*\",B2:B9,\"Mats\")";
            ws.Cells["C6"].Formula = "SUMIF(A2:A9,\"=A*\",C2:C9)";
            ws.Cells["C7"].Formula = "AVERAGEIF(A2:A9,\"=AP*\",C2:C9)";
            ws.Cells["C8"].Value = 12D;
            ws.Cells["C9"].Value = 13D;
            ws.Calculate();

            Assert.AreEqual(10.5D, ws.Cells["C4"].Value);
            Assert.AreEqual(20.5D, ws.Cells["C5"].Value);
            Assert.AreEqual(52D, ws.Cells["C6"].Value);
            Assert.AreEqual(10.5D, ws.Cells["C7"].Value);
            SaveAndCleanup(p);
        }


        [TestMethod]
        public void AverageIfsShouldNotCacheWhenValueRangeHasFormulas()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "Fruit";
            ws.Cells["B1"].Value = "Employee";
            ws.Cells["A2:A3"].Value = "Apples";
            ws.Cells["A4:A5"].Value = "Artichokes";
            ws.Cells["A6:A7"].Value = "Bananas";
            ws.Cells["A8:A9"].Value = "Carrots";
            ws.Cells["B2,B4,B8"].Value = "Mats";
            ws.Cells["B3,B5,B9"].Value = "Jan";
            ws.Cells["B6,B7"].Value = "Ossian";  // Both B6 and B7 are Ossian

            // C2, C3 have values
            ws.Cells["C2"].Value = 10D;
            ws.Cells["C3"].Value = 11D;

            // C4 and C5 have formulas that reference C2:C9 (circular reference scenario)
            ws.Cells["C4"].Formula = "AVERAGEIFS(C2:C9,A2:A9,\"=B*\",B2:B9,\"Ossian\")";
            ws.Cells["C5"].Formula = "AVERAGEIFS(C2:C9,A2:A9,\"=A*\",B2:B9,\"Mats\")";

            ws.Cells["C6"].Value = 12D;
            ws.Cells["C7"].Value = 13D;
            ws.Cells["C8"].Value = 14D;
            ws.Cells["C9"].Value = 15D;

            ws.Calculate();

            // C4 = AVERAGEIFS(C2:C9, A2:A9, "=B*", B2:B9, "Ossian")
            // Matches: A6="Bananas" (B*), B6="Ossian" -> C6=12
            //          A7="Bananas" (B*), B7="Ossian" -> C7=13
            // Average = (12+13)/2 = 12.5
            Assert.AreEqual(12.5D, ws.Cells["C4"].Value, "C4 should be 12.5");

            // C5 = AVERAGEIFS(C2:C9, A2:A9, "=A*", B2:B9, "Mats")
            // Matches: A2="Apples" (A*), B2="Mats" -> C2=10
            //          A4="Artichokes" (A*), B4="Mats" -> C4=12.5 (calculated earlier!)
            // Average = (10+12.5)/2 = 11.25
            Assert.AreEqual(11.25D, ws.Cells["C5"].Value, "C5 should be 11.25 (including calculated C4 value)");
        }
    }
}
