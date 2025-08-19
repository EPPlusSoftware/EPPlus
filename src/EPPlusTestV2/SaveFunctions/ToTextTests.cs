using FakeItEasy;
using Microsoft.SqlServer.Server;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.SaveFunctions
{
    [TestClass]
    public class ToTextTests : TestBase
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
        public void UseRowText()
        {
            _sheet.Cells["A1"].Value = "a";
            _sheet.Cells["B1"].Value = "b";
            _sheet.Cells["C1"].Value = "c";
            _sheet.Cells["A2"].Value = "d";
            _sheet.Cells["B2"].Value = "e";
            _sheet.Cells["C2"].Value = "f";
            _sheet.Cells["A3"].Value = "g";
            _sheet.Cells["B3"].Value = "h";
            _sheet.Cells["C3"].Value = "i";
            var format = new ExcelOutputTextFormat
            {
                TextQualifier = '\'',
                FirstRowIsHeader = false
            };
            format.ShouldUseRow = row =>
            {
                if (row.Contains("e"))
                {
                    return false;
                }
                return true;
            };
            var text = _sheet.Cells["A1:C3"].ToText(format);

            Assert.AreEqual("\'a\',\'b\',\'c\'" + format.EOL + "\'g\',\'h\',\'i\'", text);
        }

        [TestMethod]
        public void ToTextFixedWidth()
        {
            _sheet.Cells["A1"].Value = "Value";
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["C1"].Value = "51%";
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(5, 3, 5);
            var text = _sheet.Cells["A1:C1"].ToText(format);
            Assert.AreEqual("Value  2  51%" + format.EOL, text);
        }

        [TestMethod]
        public void ToTextLeftPaddingFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.Columns[0].PaddingType = PaddingAlignmentType.Left;
            format.Columns[1].PaddingType = PaddingAlignmentType.Left;
            var text = _sheet.Cells["A1:B1"].ToText(format);
            Assert.AreEqual("1  2   " + format.EOL, text);
        }

        [TestMethod]
        public void ToTextRightPaddingFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.Columns[0].PaddingType = PaddingAlignmentType.Right;
            format.Columns[1].PaddingType = PaddingAlignmentType.Right;
            var text = _sheet.Cells["A1:B1"].ToText(format);
            Assert.AreEqual("  1   2" + format.EOL, text);
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
            format.SetColumnLengths(3, 4, 3);
            format.ShouldUseRow = row =>
            {
                if (row.Contains("5"))
                {
                    return false;
                }
                return true;
            };
            var text = _sheet.Cells["A1:C3"].ToText(format);
            Assert.AreEqual("  1   2  3" + format.EOL + "  7   8 10" + format.EOL, text);
        }

        [TestMethod]
        public void ToTextExcludeColumnFixedWidth()
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
            format.SetColumnLengths(3, 4, 3);
            format.Columns[0].UseColumn = true;
            format.Columns[0].PaddingType = PaddingAlignmentType.Left;
            format.Columns[1].UseColumn = false;
            format.Columns[1].PaddingType = PaddingAlignmentType.Left;
            format.Columns[2].UseColumn = true;
            format.Columns[2].PaddingType = PaddingAlignmentType.Left;
            var text = _sheet.Cells["A1:C3"].ToText(format);
            Assert.AreEqual("1  3  " + format.EOL + "4  6  " + format.EOL + "7  10 " + format.EOL, text);
        }

        [TestMethod]
        public void ToTextHeaderFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = 4;
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.ExcludeHeader = true;
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual("  4   5" + format.EOL, text);
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException), "string is too long for column width")]
        public void ToTextMismatchColLengthFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = "this will throw an exception";
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            var text = _sheet.Cells["A1:B2"].ToText(format);
        }

        [TestMethod]
        public void ToTextForceWriteColLengthFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = "this will not throw an exception";
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual("  1   2" + format.EOL + "thi   5" + format.EOL, text);
        }

        [TestMethod]
        public void ToTextForceWritePositionColLengthFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = "this will not throw an exception";
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnPositions(0, 0, 3);
            format.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
            format.ReadType = FixedWidthReadType.Positions;
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual("  12" + format.EOL + "thi5" + format.EOL, text);
        }

        [TestMethod]
        public void ToTextSkipLinesBeginingFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = 4;
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.SkipLinesBeginning = 1;
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual("  4   5" + format.EOL, text);
        }

        [TestMethod]
        public void WriteFixedWidthTextFile()
        {
            using (var p = OpenTemplatePackage("Fixed3.xlsx"))
            {
                var ws = p.Workbook.Worksheets["TEST"];
                ExcelOutputTextFormatFixedWidth fw = new ExcelOutputTextFormatFixedWidth();
                fw.SetColumnPaddingAlignmentType(PaddingAlignmentType.Auto, PaddingAlignmentType.Auto, PaddingAlignmentType.Auto, PaddingAlignmentType.Auto, PaddingAlignmentType.Left, PaddingAlignmentType.Auto);
                fw.SetColumnPositions(80, 0, 16, 32, 43, 55, 62);
                fw.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
                var text = ws.Cells["A1:F2073"].ToText(fw);
                using (StreamWriter outputFile = new StreamWriter("C:\\epplusTest\\Testoutput\\NewFW2.txt"))
                {
                    outputFile.WriteLine(text);
                }
            }
        }

        [TestMethod]
        public void ToTextTrailingMinusFixedWidth()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = -4;
            _sheet.Cells["B2"].Value = 5;
            ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();
            format.SetColumnLengths(3, 4);
            format.SkipLinesBeginning = 1;
            format.UseTrailingMinus = true;
            var text = _sheet.Cells["A1:B2"].ToText(format);
            Assert.AreEqual(" 4-   5" + format.EOL, text);
        }

        [TestMethod]
        public void TransposedToText()
        {
            _sheet.Cells["A1"].Value = "Id";
            _sheet.Cells["B1"].Value = 1;
            _sheet.Cells["C1"].Value = 2;
            _sheet.Cells["D1"].Value = 3;
            _sheet.Cells["E1"].Value = 4;
            _sheet.Cells["F1"].Value = 5;
            _sheet.Cells["G1"].Value = 6;
            _sheet.Cells["A2"].Value = "Name";
            _sheet.Cells["B2"].Value = "Scott";
            _sheet.Cells["C2"].Value = "Mats";
            _sheet.Cells["D2"].Value = "Jimmy";
            _sheet.Cells["E2"].Value = "Cameron";
            _sheet.Cells["F2"].Value = "Luther";
            _sheet.Cells["G2"].Value = "Josh";

            var format = new ExcelOutputTextFormat
            {
                TextQualifier = '\'',
                DataIsTransposed = true,
            };
            var text = _sheet.Cells["A1:G2"].ToText(format);
            Assert.IsTrue(text.Contains("Luther"));
        }
    }
}

