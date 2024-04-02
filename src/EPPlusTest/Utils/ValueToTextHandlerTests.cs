using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class ValueToTextHandlerTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Init()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _sheet = null;
            _package.Dispose();
            _package = null;
        }

        [TestMethod]
        public void ShouldFormatDateWithApostrophe()
        {
            var cc = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            var date = new DateTime(2023, 12, 1);
            _sheet.Cells["A1"].Value = date;
            _sheet.Cells["A1"].Style.Numberformat.Format = "MMMM \\'yy";
            var text = _sheet.Cells["A1"].Text;
            Thread.CurrentThread.CurrentCulture = cc;
            Assert.AreEqual("December '23", text);
        }
    }
}
