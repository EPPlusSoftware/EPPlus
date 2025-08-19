using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public  class HstackTests
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
        public void BasicTest1()
        {
            _sheet.Cells["A1"].Value = 1;
            _sheet.Cells["B1"].Value = 2;
            _sheet.Cells["A2"].Value = 3;
            _sheet.Cells["B2"].Value = 4;
            _sheet.Cells["A3"].Value = 5;
            _sheet.Cells["B3"].Value = 6;

            _sheet.Cells["D1"].Value = 100;
            _sheet.Cells["D2"].Value = 200;

            _sheet.Cells["F1"].Value = 1000;
            _sheet.Cells["G1"].Value = 1001;

            _sheet.Cells["A6"].Formula = "HSTACK(A1:B3,D1:D2,F1:G1)";
            _sheet.Cells["A6"].Calculate();

            Assert.AreEqual(1, _sheet.Cells["A6"].Value);
            Assert.AreEqual(6, _sheet.Cells["B8"].Value);
            Assert.AreEqual(100, _sheet.Cells["C6"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["C8"].Value);
            Assert.AreEqual(1000, _sheet.Cells["D6"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["D7"].Value);
            Assert.AreEqual(1001, _sheet.Cells["E6"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["E7"].Value);
            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells["E8"].Value);
        }
    }
}
