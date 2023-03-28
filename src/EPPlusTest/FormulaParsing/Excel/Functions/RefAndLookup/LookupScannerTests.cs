using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class LookupScannerTests
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
        public void ShouldFindExactMatch_StartAtFirst_Vertical_1()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 6;
            _sheet.Cells[5, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A5"]);
            var scanner = new XlookupScanner(4, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(2, ix);
        }

        [TestMethod]
        public void ShouldReturnNotFoundExactMatch_StartAtFirst_Vertical()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 6;
            _sheet.Cells[5, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A5"]);
            var scanner = new XlookupScanner(8, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatch);
            var ix = scanner.FindIndex();
            Assert.AreEqual(-1, ix);
        }

        [TestMethod]
        public void ShouldFindExactMatch_StartAtFirst_Vertical_2()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 9;
            _sheet.Cells[5, 1].Value = 6;
            _sheet.Cells[6, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A6"]);
            var scanner = new XlookupScanner(9, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(3, ix);
        }

        [TestMethod]
        public void ShouldFindClosestAbove_StartAtFirst_Vertical()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 6;
            _sheet.Cells[5, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A5"]);
            var scanner = new XlookupScanner(8, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(4, ix);
        }

        [TestMethod]
        public void ShouldFindClosestBelow_StartAtFirst_Vertical()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 6;
            _sheet.Cells[5, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A5"]);
            var scanner = new XlookupScanner(8, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchReturnNextSmaller);
            var ix = scanner.FindIndex();
            Assert.AreEqual(3, ix);
        }

        [TestMethod]
        public void ShouldFindExactMatch_StartAtLast_Vertical_1()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 9;
            _sheet.Cells[4, 1].Value = 6;
            _sheet.Cells[5, 1].Value = 9;
            _sheet.Cells[6, 1].Value = 4;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A6"]);
            var scanner = new XlookupScanner(9, ri, LookupSearchMode.ReverseStartingAtLast, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(4, ix);
        }

        [TestMethod]
        public void ShouldFindExactMatch_StartAtLast_Vertical_2()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[2, 1].Value = 15;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[4, 1].Value = 9;
            _sheet.Cells[5, 1].Value = 6;
            _sheet.Cells[6, 1].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:A6"]);
            var scanner = new XlookupScanner(9, ri, LookupSearchMode.ReverseStartingAtLast, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(5, ix);
        }

        [TestMethod]
        public void ShouldFindExactMatch_StartAtFirst_Horizontal_1()
        {
            _sheet.Cells[1, 1].Value = 10;
            _sheet.Cells[1, 2].Value = 15;
            _sheet.Cells[1, 3].Value = 4;
            _sheet.Cells[1, 4].Value = 6;
            _sheet.Cells[1, 5].Value = 9;

            var ri = new RangeInfo(_sheet, _sheet.Cells["A1:E1"]);
            var scanner = new XlookupScanner(4, ri, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchReturnNextLarger);
            var ix = scanner.FindIndex();
            Assert.AreEqual(2, ix);
        }
    }
}
