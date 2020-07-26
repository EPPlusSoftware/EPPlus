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
    }
}
