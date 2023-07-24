using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class CountIfTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void CountIfShouldCalculateOneCriteria()
        {
            _sheet.Cells["A1"].Value = 10;
            _sheet.Cells["A2"].Value = 11;
            _sheet.Cells["A3"].Value = 12;
            _sheet.Cells["A4"].Value = 13;
            _sheet.Cells["A5"].Value = 14;
            _sheet.Cells["A6"].Value = 15;
            _sheet.Cells["B1"].Formula = "COUNTIF(A1:A6,\">13\")";
            _sheet.Calculate();
            Assert.AreEqual(2d, _sheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void CountIfShouldCalculateRangeCriteria()
        {
            _sheet.Cells["A1"].Value = 10;
            _sheet.Cells["A2"].Value = 11;
            _sheet.Cells["A3"].Value = 12;
            _sheet.Cells["A4"].Value = 12;
            _sheet.Cells["A5"].Value = 14;
            _sheet.Cells["A6"].Value = 15;

            _sheet.Cells["B1"].Value = 10;
            _sheet.Cells["B2"].Value = 12;

            _sheet.Cells["C1"].Formula = "SUM(COUNTIF(A1:A6,B1:B2))";
            _sheet.Calculate();
            Assert.AreEqual(3m, _sheet.Cells["C1"].Value);
        }
    }
}
