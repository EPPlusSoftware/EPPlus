using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class FrequencyTests
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
        public void BasicFrequencyTest1()
        {
            _sheet.Cells["A1"].Value = 1d;
            _sheet.Cells["A2"].Value = 1d;
            _sheet.Cells["A3"].Value = 3d;
            _sheet.Cells["A4"].Value = 5d;
            _sheet.Cells["A5"].Value = 7d;
            _sheet.Cells["A6"].Value = 4d;
            _sheet.Cells["B1"].Value = 2d;
            _sheet.Cells["B2"].Value = 4d;
            _sheet.Cells["B3"].Value = 6d;
            _sheet.Cells["A8"].Formula = "FREQUENCY(A1:A6,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(2, _sheet.Cells["A8"].Value);
            Assert.AreEqual(2, _sheet.Cells["A9"].Value);

        }

        [TestMethod]
        public void UnsortedBinsArray()
        {
            _sheet.Cells["A1"].Value = 1d;
            _sheet.Cells["A2"].Value = 1d;
            _sheet.Cells["A3"].Value = 3d;
            _sheet.Cells["A4"].Value = 5d;
            _sheet.Cells["A5"].Value = 7d;
            _sheet.Cells["A6"].Value = 4d;
            _sheet.Cells["B1"].Value = 6d;
            _sheet.Cells["B2"].Value = 2d;
            _sheet.Cells["B3"].Value = 4d;
            _sheet.Cells["A8"].Formula = "FREQUENCY(A1:A6,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells["A8"].Value);
            Assert.AreEqual(2, _sheet.Cells["A9"].Value);
        }

        [TestMethod]
        public void EmptyBinsArray()
        {
            _sheet.Cells["A1"].Value = 1d;
            _sheet.Cells["A2"].Value = 1d;
            _sheet.Cells["A3"].Value = 3d;
            _sheet.Cells["A4"].Value = 5d;
            _sheet.Cells["A5"].Value = 7d;
            _sheet.Cells["A6"].Value = 4d;
            _sheet.Cells["A8"].Formula = "FREQUENCY(A1:A6,B1:B3)";
            _sheet.Calculate();
            Assert.AreEqual(0, _sheet.Cells["A8"].Value);
            Assert.AreEqual(6, _sheet.Cells["A9"].Value);
        }

        [TestMethod]
        public void NonDoublesArray()
        {
            _sheet.Cells["A1"].Value = "a";
            _sheet.Cells["A2"].Value = "b";
            _sheet.Cells["A3"].Value = "c";
            _sheet.Cells["A4"].Value = "c";
            _sheet.Cells["A5"].Value = "d";
            _sheet.Cells["A6"].Value = "e";
            _sheet.Cells["B1"].Value = "b";
            _sheet.Cells["B2"].Value = "d";
            _sheet.Cells["A8"].Formula = "FREQUENCY(A1:A6,B1:B2)";
            _sheet.Calculate();
            Assert.AreEqual(0, _sheet.Cells["A8"].Value);
            Assert.AreEqual(0, _sheet.Cells["A9"].Value);
        }
    }
}
