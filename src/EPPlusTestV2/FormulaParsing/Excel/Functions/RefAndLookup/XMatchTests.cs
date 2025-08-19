using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class XMatchTests
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
        public void BasicTest()
        {
            _sheet.Cells[1, 1].Value = "Apple";
            _sheet.Cells[2, 1].Value = "Grape";
            _sheet.Cells[3, 1].Value = "Pear";
            _sheet.Cells[4, 1].Value = "Banana";
            _sheet.Cells[5, 1].Value = "Cherry";

            _sheet.Cells[6, 1].Formula = "XMATCH(\"Grape\",A1:A5)";
            _sheet.Calculate();

            Assert.AreEqual(2, _sheet.Cells[6, 1].Value, "Exact match not detected");

            _sheet.Cells[6, 1].Formula = "XMATCH(\"Grpe\",A1:A5)";
            _sheet.Calculate();

            Assert.AreEqual(ErrorValues.NAError, _sheet.Cells[6, 1].Value, "Value was not NA error as expected");
        }

        [TestMethod]
        public void FromLastSearchMode()
        {
            _sheet.Cells[1, 1].Value = "Apple";
            _sheet.Cells[2, 1].Value = "Grape";
            _sheet.Cells[3, 1].Value = "Pear";
            _sheet.Cells[4, 1].Value = "Pear";
            _sheet.Cells[5, 1].Value = "Cherry";

            _sheet.Cells[6, 1].Formula = "XMATCH(\"Pear\",A1:A5,0,-1)";
            _sheet.Calculate();

            Assert.AreEqual(4, _sheet.Cells[6, 1].Value, "Exact match not detected");

        }

        [TestMethod]
        public void BinarySearchAsc()
        {
            _sheet.Cells[1, 1].Value = "Apple";
            _sheet.Cells[2, 1].Value = "Banana";
            _sheet.Cells[3, 1].Value = "Cherry";
            _sheet.Cells[4, 1].Value = "Grape";
            _sheet.Cells[5, 1].Value = "Pear";

            // return next smaller
            _sheet.Cells[6, 1].Formula = "XMATCH(\"Ara\",A1:A5,-1,2)";
            _sheet.Calculate();
            Assert.AreEqual(1, _sheet.Cells[6, 1].Value, "Error when returning next smaller");

            // return next larger
            _sheet.Cells[6, 1].Formula = "XMATCH(\"Cool\",A1:A5,1,2)";
            _sheet.Calculate();
            Assert.AreEqual(4, _sheet.Cells[6, 1].Value, "Error when returning next larger");

        }

        [TestMethod]
        public void BinarySearchDesc()
        {
            _sheet.Cells[1, 1].Value = "Pear";
            _sheet.Cells[2, 1].Value = "Grape";
            _sheet.Cells[3, 1].Value = "Cherry";
            _sheet.Cells[4, 1].Value = "Banana";
            _sheet.Cells[5, 1].Value = "Apple";

            // return next smaller
            _sheet.Cells[6, 1].Formula = "XMATCH(\"Ara\",A1:A5,-1,-2)";
            _sheet.Calculate();
            Assert.AreEqual(4, _sheet.Cells[6, 1].Value, "Error when returning next smaller");

            // return next larger
            _sheet.Cells[6, 1].Formula = "XMATCH(\"Ara\",A1:A5,1,-2)";
            _sheet.Calculate();
            Assert.AreEqual(5, _sheet.Cells[6, 1].Value, "Error when returning next larger");

        }
    }
}
