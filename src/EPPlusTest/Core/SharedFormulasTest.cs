using EPPlusTest.Drawing.Chart.Styling;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;


namespace OfficeOpenXml.FormulaParsing
{
    [TestClass]
    public class SharedFormulasTest
    {
        private static ExcelPackage _package;
        private static ExcelWorksheet _ws;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _package = new ExcelPackage();
            _ws = _package.Workbook.Worksheets.Add("Sheet1");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _package.Dispose();
        }
        [TestMethod]
        public void SharedFormulasShouldNotEffectFullColumn()
        {
            var f=new SharedFormula() { Index = 0, Formula = "SUM(C:D)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

            var fA1= f.GetFormula(1, 1, "sheet1");
            var fA2 = f.GetFormula(2, 1, "sheet1");
            var fB1 = f.GetFormula(1, 2, "sheet1");
            var fB2 = f.GetFormula(2, 2, "sheet1");

            Assert.AreEqual("SUM(C:D)", fA1);
            Assert.AreEqual("SUM(C:D)", fA2);
            Assert.AreEqual("SUM(D:E)", fB1);
            Assert.AreEqual("SUM(D:E)", fB2);
        }
        [TestMethod]
        public void SharedFormulasShouldNotEffectFullRow()
        {
            var f = new SharedFormula() { Index = 0, Formula = "SUM(3:4)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

            var fA1 = f.GetFormula(1, 1, "sheet1");
            var fA2 = f.GetFormula(2, 1, "sheet1");
            var fB1 = f.GetFormula(1, 2, "sheet1");
            var fB2 = f.GetFormula(2, 2, "sheet1");

            Assert.AreEqual("SUM(3:4)", fA1);
            Assert.AreEqual("SUM(4:5)", fA2);
            Assert.AreEqual("SUM(3:4)", fB1);
            Assert.AreEqual("SUM(4:5)", fB2);
        }
        [TestMethod]
        public void SharedFormulasShouldNotEffectFullSheet()
        {
            var f = new SharedFormula() { Index = 0, Formula = "SUM(A:XFD)", Address = "A1:B2", StartRow = 1, StartCol = 1 };

            var fA1 = f.GetFormula(1, 1, "sheet1");
            var fA2 = f.GetFormula(2, 1, "sheet1");
            var fB1 = f.GetFormula(1, 2, "sheet1");
            var fB2 = f.GetFormula(2, 2, "sheet1");

            Assert.AreEqual("SUM(A:XFD)", fA1);
            Assert.AreEqual("SUM(A:XFD)", fA2);
            Assert.AreEqual("SUM(B:#REF!)", fB1);
            Assert.AreEqual("SUM(B:#REF!)", fB2);
        }
    }
}
