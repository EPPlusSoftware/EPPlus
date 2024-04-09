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
