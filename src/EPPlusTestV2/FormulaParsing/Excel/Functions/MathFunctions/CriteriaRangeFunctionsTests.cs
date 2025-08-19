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
    }
}
