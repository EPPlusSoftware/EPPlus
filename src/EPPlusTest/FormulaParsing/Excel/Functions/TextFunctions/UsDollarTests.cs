using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Text
{
    [TestClass]
    public class UsDollarTests : TestBase
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _package;
        private CultureInfo _originalCulture;

        [TestInitialize]
        public void TestInitialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
            SwitchToCulture("en-US");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            SwitchBackToCurrentCulture();
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldFormatPositiveNumberWithDefaultDecimals()
        {
            _sheet.Cells["A1"].Formula = "USDOLLAR(1234.567)";
            _sheet.Calculate();
            Assert.AreEqual("$1,234.57", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldFormatNegativeNumberWithParentheses()
        {
            _sheet.Cells["A1"].Formula = "USDOLLAR(-1234.567)";
            _sheet.Calculate();
            // en-US currency format renders negatives as ($1,234.57)
            Assert.AreEqual("($1,234.57)", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldHonorExplicitDecimals()
        {
            _sheet.Cells["A1"].Formula = "USDOLLAR(1234.5678,3)";
            _sheet.Calculate();
            Assert.AreEqual("$1,234.568", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldFormatWithZeroDecimals()
        {
            _sheet.Cells["A1"].Formula = "USDOLLAR(1234.5,0)";
            _sheet.Calculate();
            Assert.AreEqual("$1,235", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldRoundToNegativeDecimals()
        {
            // -2 decimals rounds to nearest hundred
            _sheet.Cells["A1"].Formula = "USDOLLAR(1234.567,-2)";
            _sheet.Calculate();
            Assert.AreEqual("$1,200", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldHandleZero()
        {
            _sheet.Cells["A1"].Formula = "USDOLLAR(0)";
            _sheet.Calculate();
            Assert.AreEqual("$0.00", _sheet.Cells["A1"].Value);
        }
    }
}