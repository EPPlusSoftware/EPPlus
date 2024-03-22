using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.SaveFunctions
{
    [TestClass]
    public class ToTextTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ToTextTextDefault()
        {
            _sheet.Cells["A1"].Value = "h1";
            _sheet.Cells["B1"].Value = "h2";
            var text = _sheet.Cells["A1:B1"].ToText();
            Assert.AreEqual("h1,h2", text);
        }

        [TestMethod]
        public void ToTextTextMultilines()
        {
            _sheet.Cells["A1"].Value = "h1";
            _sheet.Cells["B1"].Value = "h2";
            _sheet.Cells["A2"].Value = 1;
            _sheet.Cells["B2"].Value = 2;
            var text = _sheet.Cells["A1:B2"].ToText();
            Assert.AreEqual("h1,h2" + Environment.NewLine + "1,2", text);
        }

        [TestMethod]
        public void ToTextTextTextQualifier()
        {
            _sheet.Cells["A1"].Value = "h1";
            _sheet.Cells["B1"].Value = "h2";
            _sheet.Cells["A2"].Value = 1;
            _sheet.Cells["B2"].Value = 2;
            var format = new ExcelOutputTextFormat
            {
                TextQualifier = '\''
            };
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual("'h1','h2'" + Environment.NewLine + "1,2", text);
        }
        [TestMethod]
        public void ToTextTextQualifierWithNumericContaingSeparator()
        {
            _sheet.Cells["A10"].Value = "h1";
            _sheet.Cells["B10"].Value = "h2";
            _sheet.Cells["A11"].Value = 1;
            _sheet.Cells["B11"].Value = 2;
            _sheet.Cells["A11:B11"].Style.Numberformat.Format = "#,##0.00";
            var format = new ExcelOutputTextFormat
            {
                TextQualifier = '\"',
                DecimalSeparator = ",",
                UseCellFormat = true,
                Culture = new System.Globalization.CultureInfo("sv-SE")
            };
            var text = _sheet.Cells["A10:B11"].ToText(format);
            Assert.AreEqual("\"h1\",\"h2\"" + Environment.NewLine + "\"1,00\",\"2,00\"", text);
        }

        [TestMethod]
        public void ToTextTextIgnoreHeaders()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            var format = new ExcelOutputTextFormat
            {
                TextQualifier = '\'',
                FirstRowIsHeader = false
            };
            var text = _sheet.Cells["A1:B1"].ToText(format);
            Assert.AreEqual("1,2", text);
        }

        //[TestMethod]
        //public void ToTextFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    var text = _sheet.Cells["A1:B1"].ToText(format);
        //    Assert.AreEqual("1  2   " + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextLeftPaddingFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    format.PaddingType = SpacePaddingType.Left;
        //    var text = _sheet.Cells["A1:B1"].ToText(format);
        //    Assert.AreEqual("  1   2" + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextExcludeRowFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["C1"].Value = 3;

        //    _sheet.Cells["A2"].Value = 4;
        //    _sheet.Cells["B2"].Value = 5;
        //    _sheet.Cells["C2"].Value = 6;

        //    _sheet.Cells["A3"].Value = 7;
        //    _sheet.Cells["B3"].Value = 8;
        //    _sheet.Cells["C3"].Value = 10;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4, 3};
        //    format.PaddingType = SpacePaddingType.Left;
        //    format.ShouldUseRow = row =>
        //    {
        //        if (row.Contains("5"))
        //        {
        //            return false;
        //        }
        //        return true;
        //    };
        //    var text = _sheet.Cells["A1:C3"].ToText(format);
        //    Assert.AreEqual("  1   2  3" + format.EOL + "  7   8 10" + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextExcludeColumnFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["C1"].Value = 3;

        //    _sheet.Cells["A2"].Value = 4;
        //    _sheet.Cells["B2"].Value = 5;
        //    _sheet.Cells["C2"].Value = 6;

        //    _sheet.Cells["A3"].Value = 7;
        //    _sheet.Cells["B3"].Value = 8;
        //    _sheet.Cells["C3"].Value = 10;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4, 3 };
        //    format.UseColumns =  new bool[] { true, false, true };
        //    var text = _sheet.Cells["A1:C3"].ToText(format);
        //    Assert.AreEqual("1  3  " + format.EOL + "4  6  " + format.EOL + "7  10 " + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextHeaderFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["A2"].Value = 4;
        //    _sheet.Cells["B2"].Value = 5;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    format.ExcludeHeader = true;
        //    var text = _sheet.Cells["A1:B2"].ToText(format);
        //    Assert.AreEqual("4  5   " + format.EOL, text);
        //}

        //[TestMethod]
        //[ExpectedException(typeof(FormatException), "string is too long for column width")]
        //public void ToTextMismatchColLengthFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["A2"].Value = "this will throw an exception";
        //    _sheet.Cells["B2"].Value = 5;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    var text = _sheet.Cells["A1:B2"].ToText(format);
        //}

        //[TestMethod]
        //public void ToTextForceWriteColLengthFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["A2"].Value = "this will not throw an exception";
        //    _sheet.Cells["B2"].Value = 5;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    format.ForceWrite = true;
        //    var text = _sheet.Cells["A1:B2"].ToText(format);
        //    Assert.AreEqual("1  2   " + format.EOL + "thi5   " + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextForceWritePositionColLengthFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["A2"].Value = "this will not throw an exception";
        //    _sheet.Cells["B2"].Value = 5;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 0, 3 };
        //    format.ForceWrite = true;
        //    format.ReadStartPosition = FixedWidthReadType.Positions;
        //    var text = _sheet.Cells["A1:B2"].ToText(format);
        //    Assert.AreEqual("1  2" + format.EOL + "thi5" + format.EOL, text);
        //}

        //[TestMethod]
        //public void ToTextSkipLinesBeginingFixedWidth()
        //{
        //    _sheet.Cells["A1"].Value = 1;
        //    _sheet.Cells["B1"].Value = 2;
        //    _sheet.Cells["A2"].Value = 4;
        //    _sheet.Cells["B2"].Value = 5;
        //    ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
        //    format.ColumnLengths = new int[] { 3, 4 };
        //    format.SkipLinesBeginning = 1;
        //    var text = _sheet.Cells["A1:B2"].ToText(format);
        //    Assert.AreEqual("4  5   " + format.EOL, text);
        //}

        /* make creating list with excelTextFormatColumns easier.
         * detect data type when reading/saving and set paddingalignment based on that.
         * when reading from position, make it possible to read length of last column for reading and saving file
         * make so we can read/save empty or shorter lines than expected
         * force read row om den är för kort// inte passa spec Loadfromfixedwidthtext(då läser vi till raden tar slut så vi inte skriver utanför length)
         * shouldUseRow CSV
         * exclude column CSV reading and saving
         */
    }
}
