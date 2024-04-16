using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;


namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public  class MMultTests : TestBase
    {
        [TestMethod]
        public void MMultTest()
        {
            using var p = OpenPackage("MMultTest.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 5;
            ws.Cells["B1"].Value = 6;
            ws.Cells["C1"].Value = 7;
            ws.Cells["A2"].Value = 3;
            ws.Cells["B2"].Value = 4;
            ws.Cells["C2"].Value = 2;

            ws.Cells["A4"].Value = 9;
            ws.Cells["A5"].Value = 1;
            ws.Cells["A6"].Value = 3;
            ws.Cells["B4"].Value = 8;
            ws.Cells["B5"].Value = 1;
            ws.Cells["B6"].Value = 3;

            ws.Cells["E1"].Formula = "MMULT(A1:C2,A4:B6)";
            ws.Calculate();
            Assert.AreEqual(72d, ws.Cells["E1"].Value);
            Assert.AreEqual(67d, ws.Cells["F1"].Value);
            Assert.AreEqual(37d, ws.Cells["E2"].Value);
            Assert.AreEqual(34d, ws.Cells["F2"].Value);

        }
    }
}
