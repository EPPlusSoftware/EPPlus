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

        [TestMethod]
        public void ToTextFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.ColumnLengths = new int[] { 3, 4 };
            format.FirstRowIsHeader = false;
            var text = _sheet.Cells["A1:B1"].ToText(format);
            Assert.AreEqual("1  2   ", text);
        }

        [TestMethod]
        public void ToTextLeftPaddingFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.ColumnLengths = new int[] { 3, 4 };
            format.FirstRowIsHeader = false;
            format.PaddingType = SpacePaddingType.Left;
            var text = _sheet.Cells["A1:B1"].ToText(format);
            Assert.AreEqual("  1   2", text);
        }

        [TestMethod]
        public void ToTextExcludeRowFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["C1"].Value = 3;

            _sheet.Cells["A2"].Value = 4;
            _sheet.Cells["B2"].Value = 5;
            _sheet.Cells["C2"].Value = 6;

            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B3"].Value = 8;
            _sheet.Cells["C3"].Value = 10;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.ColumnLengths = new int[] { 3, 4, 3};
            format.FirstRowIsHeader = false;
            format.PaddingType = SpacePaddingType.Left;
            format.ShouldUseRow = row =>
            {
                if (row.Contains("5"))
                {
                    return false;
                }
                return true;
            };
            var text = _sheet.Cells["A1:C3"].ToText(format);
            Assert.AreEqual("  1   2  3\r\n  7   8 10\r\n", text);
        }

        [TestMethod]
        public void ToTextExcludeRowFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["C1"].Value = 3;

            _sheet.Cells["A2"].Value = 4;
            _sheet.Cells["B2"].Value = 5;
            _sheet.Cells["C2"].Value = 6;

            _sheet.Cells["A3"].Value = 7;
            _sheet.Cells["B3"].Value = 8;
            _sheet.Cells["C3"].Value = 10;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.ColumnLengths = new int[] { 3, 4, 3 };
            format.FirstRowIsHeader = false;
            format.PaddingType = SpacePaddingType.Left;
            format.ShouldUseRow = row =>
            {
                if (row.Contains("5"))
                {
                    return false;
                }
                return true;
            };
            var text = _sheet.Cells["A1:C3"].ToText(format);
            Assert.AreEqual("  1   2  3\r\n  7   8 10\r\n", text);
        }

    }
}
