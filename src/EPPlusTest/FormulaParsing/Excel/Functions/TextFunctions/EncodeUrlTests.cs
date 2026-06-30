using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Text
{
    [TestClass]
    public class EncodeUrlTests
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
        public void ShouldEncodeSpacesAsPercent20()
        {
            _sheet.Cells["A1"].Formula = "ENCODEURL(\"hello world\")";
            _sheet.Calculate();
            Assert.AreEqual("hello%20world", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldEncodeSpecialCharacters()
        {
            _sheet.Cells["A1"].Formula = "ENCODEURL(\"a&b=c?d\")";
            _sheet.Calculate();
            Assert.AreEqual("a%26b%3Dc%3Fd", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldNotEncodeUnreservedAscii()
        {
            _sheet.Cells["A1"].Formula = "ENCODEURL(\"abc-XYZ_123.~\")";
            _sheet.Calculate();
            // Letters, digits, hyphen, underscore, period and tilde are unreserved per RFC 3986.
            Assert.AreEqual("abc-XYZ_123.~", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldEncodeUnicodeAsUtf8()
        {
            _sheet.Cells["A1"].Formula = "ENCODEURL(\"\u00e5\u00e4\u00f6\")";
            _sheet.Calculate();
            // å = C3 A5, ä = C3 A4, ö = C3 B6 in UTF-8
            Assert.AreEqual("%C3%A5%C3%A4%C3%B6", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldReturnEmptyStringForEmptyInput()
        {
            _sheet.Cells["A1"].Formula = "ENCODEURL(\"\")";
            _sheet.Calculate();
            Assert.AreEqual("", _sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldEncodeCellReference()
        {
            _sheet.Cells["B1"].Value = "test value";
            _sheet.Cells["A1"].Formula = "ENCODEURL(B1)";
            _sheet.Calculate();
            Assert.AreEqual("test%20value", _sheet.Cells["A1"].Value);
        }
    }
}