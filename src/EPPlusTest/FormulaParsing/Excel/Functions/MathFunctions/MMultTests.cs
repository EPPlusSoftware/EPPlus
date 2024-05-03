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

            ws.Cells["A10"].Value = 1;
            ws.Cells["B10"].Value = 76;
            ws.Cells["C10"].Value = 435;
            ws.Cells["D10"].Value = 987;

            ws.Cells["A11"].Value = 98;
            ws.Cells["B11"].Value = 56;
            ws.Cells["C11"].Value = 47;
            ws.Cells["D11"].Value = 8;

            ws.Cells["A12"].Value = 9;
            ws.Cells["B12"].Value = 56;
            ws.Cells["C12"].Value = 64;
            ws.Cells["D12"].Value = 8;

            ws.Cells["A13"].Value = 12;
            ws.Cells["B13"].Value = 4;
            ws.Cells["C13"].Value = 56;
            ws.Cells["D13"].Value = 7;

            ws.Cells["F10"].Value = 5;
            ws.Cells["F11"].Value = 2;
            ws.Cells["F12"].Value = 2;
            ws.Cells["F13"].Value = 6;

            ws.Cells["I10"].Formula = "MMULT(A10:D13,F10:F13)";
            ws.Calculate();
            Assert.AreEqual(6949d, ws.Cells["I10"].Value);
            Assert.AreEqual(744d, ws.Cells["I11"].Value);
            Assert.AreEqual(333d, ws.Cells["I12"].Value);
            Assert.AreEqual(222d, ws.Cells["I13"].Value);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void FaultyMatrix()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 5;
            ws.Cells["B1"].Value = 6;
            ws.Cells["C1"].Value = 7;
            ws.Cells["A2"].Value = 3;
            ws.Cells["C2"].Value = 2;

            ws.Cells["A4"].Value = 9;
            ws.Cells["A5"].Value = 1;
            ws.Cells["A6"].Value = 3;
            ws.Cells["B4"].Value = 8;
            ws.Cells["B5"].Value = 1;
            ws.Cells["B6"].Value = 3;
            ws.Cells["E1"].Formula = "MMULT(A1:C2,A4:B6)";
            ws.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E1"].Value);
        }
    }
}
