using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Text
{
    [TestClass]
    public class CodeTests
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;

        [TestInitialize]
        public void TestInitialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldReturnCodeForUppercaseA()
        {
            _sheet.Cells["A1"].Formula = "CODE(\"A\")";
            _sheet.Calculate();
            Assert.AreEqual(65d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnCodeForLowercaseA()
        {
            _sheet.Cells["A1"].Formula = "CODE(\"a\")";
            _sheet.Calculate();
            Assert.AreEqual(97d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnCodeForFirstCharacterOnly()
        {
            _sheet.Cells["A1"].Formula = "CODE(\"Hello\")";
            _sheet.Calculate();
            // H = 72
            Assert.AreEqual(72d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnCodeForDigit()
        {
            _sheet.Cells["A1"].Formula = "CODE(\"0\")";
            _sheet.Calculate();
            Assert.AreEqual(48d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnCodeForUnicodeCharacter()
        {
            // Swedish å = U+00E5 = 229
            _sheet.Cells["A1"].Formula = "CODE(\"\u00e5\")";
            _sheet.Calculate();
            Assert.AreEqual(229d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnCodeForSpace()
        {
            _sheet.Cells["A1"].Formula = "CODE(\" \")";
            _sheet.Calculate();
            Assert.AreEqual(32d, _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnValueErrorForEmptyString()
        {
            _sheet.Cells["A1"].Formula = "CODE(\"\")";
            _sheet.Calculate();
            var err = _sheet.Cells["A1"].Value as ExcelErrorValue;
            Assert.IsNotNull(err);
            Assert.AreEqual(eErrorType.Value, err.Type);
        }

        [TestMethod]
        public void ShouldReturnCodeFromCellReference()
        {
            _sheet.Cells["B1"].Value = "Z";
            _sheet.Cells["A1"].Formula = "CODE(B1)";
            _sheet.Calculate();
            Assert.AreEqual(90d, _sheet.Cells["A1"].Value);
        }
    }
}