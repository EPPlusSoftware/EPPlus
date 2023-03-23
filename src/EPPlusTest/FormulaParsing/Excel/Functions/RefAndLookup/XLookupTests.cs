using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.XlookupUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class XLookupTests
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

        [DataTestMethod]
        [DataRow("Brazil", "+55")]
        [DataRow("Sweden", "#N/A")]
        [DataRow("Sweden", "Not found", "Not found")]
        public void BasicTest(string country, string expected, string notFoundText = "")
        {
            _sheet.Cells[1, 1].Value = "China";
            _sheet.Cells[1, 2].Value = "CN";
            _sheet.Cells[1, 3].Value = "+86";
            _sheet.Cells[2, 1].Value = "India";
            _sheet.Cells[2, 2].Value = "IN";
            _sheet.Cells[2, 3].Value = "+91";
            _sheet.Cells[3, 1].Value = "United States";
            _sheet.Cells[3, 2].Value = "US";
            _sheet.Cells[3, 3].Value = "+1";
            _sheet.Cells[4, 1].Value = "Indonesia";
            _sheet.Cells[4, 2].Value = "ID";
            _sheet.Cells[4, 3].Value = "+62";
            _sheet.Cells[5, 1].Value = "Brazil";
            _sheet.Cells[5, 2].Value = "BR";
            _sheet.Cells[5, 3].Value = "+55";
            _sheet.Cells[6, 1].Value = "Pakistan";
            _sheet.Cells[6, 2].Value = "PK";
            _sheet.Cells[6, 3].Value = "+92";

            _sheet.Cells["E2"].Value = country;
            if (string.IsNullOrEmpty(notFoundText))
            {
                _sheet.Cells["F2"].Formula = "XLOOKUP(E2,A2:A11,C2:C11)";
            }
            else
            {
                _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A2:A11,C2:C11,\"{notFoundText}\")";
            }

            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [DataTestMethod]
        [DataRow("Brazil", "+55", "Not found", 0)]
        [DataRow("Brasil", "Not found", "Not found", 0)]
        [DataRow("Bazil", "#N/A", null, 0)]
        [DataRow("Bazil", "+880", "Not found", -1)]
        [DataRow("Bsazil", "+62", "Not found", 1)]
        public void SearchModeTest(string country, string expected, string notFoundText = null, int searchMode = 0)
        {
            _sheet.Cells[1, 1].Value = "China";
            _sheet.Cells[1, 2].Value = "CN";
            _sheet.Cells[1, 3].Value = "+86";
            _sheet.Cells[2, 1].Value = "India";
            _sheet.Cells[2, 2].Value = "IN";
            _sheet.Cells[2, 3].Value = "+91";
            _sheet.Cells[3, 1].Value = "United States";
            _sheet.Cells[3, 2].Value = "US";
            _sheet.Cells[3, 3].Value = "+1";
            _sheet.Cells[4, 1].Value = "Indonesia";
            _sheet.Cells[4, 2].Value = "ID";
            _sheet.Cells[4, 3].Value = "+62";
            _sheet.Cells[5, 1].Value = "Brazil";
            _sheet.Cells[5, 2].Value = "BR";
            _sheet.Cells[5, 3].Value = "+55";
            _sheet.Cells[6, 1].Value = "Pakistan";
            _sheet.Cells[6, 2].Value = "PK";
            _sheet.Cells[6, 3].Value = "+92";
            _sheet.Cells[6, 1].Value = "Bangladesh";
            _sheet.Cells[6, 2].Value = "BN";
            _sheet.Cells[6, 3].Value = "+880";

            _sheet.Cells["E2"].Value = country;
            if (notFoundText == null)
            {
                _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A2:A11,C2:C11,, {searchMode})";
            }
            else
            {
                _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A2:A11,C2:C11,\"{notFoundText}\", {searchMode})";
            }

            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [DataTestMethod]
        [DataRow("raz", "+55")]
        [DataRow("United", "+1")]
        [DataRow("desh", "+880")]
        public void WildcardTest(string lookupValue, string expected)
        {
            _sheet.Cells[1, 1].Value = "China";
            _sheet.Cells[1, 2].Value = "CN";
            _sheet.Cells[1, 3].Value = "+86";
            _sheet.Cells[2, 1].Value = "India";
            _sheet.Cells[2, 2].Value = "IN";
            _sheet.Cells[2, 3].Value = "+91";
            _sheet.Cells[3, 1].Value = "United States";
            _sheet.Cells[3, 2].Value = "US";
            _sheet.Cells[3, 3].Value = "+1";
            _sheet.Cells[4, 1].Value = "Indonesia";
            _sheet.Cells[4, 2].Value = "ID";
            _sheet.Cells[4, 3].Value = "+62";
            _sheet.Cells[5, 1].Value = "Brazil";
            _sheet.Cells[5, 2].Value = "BR";
            _sheet.Cells[5, 3].Value = "+55";
            _sheet.Cells[6, 1].Value = "Pakistan";
            _sheet.Cells[6, 2].Value = "PK";
            _sheet.Cells[6, 3].Value = "+92";
            _sheet.Cells[6, 1].Value = "Bangladesh";
            _sheet.Cells[6, 2].Value = "BN";
            _sheet.Cells[6, 3].Value = "+880";

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(\"*\"&E2&\"*\",A2:A11,C2:C11, \"Not found\", 2)";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [DataTestMethod]
        [DataRow("A", "1", 1)]
        [DataRow("A", "3", -1)]
        public void ReverseSearchTest1(string lookupValue, string expected, int searchMode)
        {
            _sheet.Cells[1, 1].Value = "A";
            _sheet.Cells[1, 2].Value = "1";
            _sheet.Cells[2, 1].Value = "B";
            _sheet.Cells[2, 2].Value = "2";
            _sheet.Cells[3, 1].Value = "A";
            _sheet.Cells[3, 2].Value = "3";

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(\"A\",A1:A3,B1:B3, \"Not found\", 0, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [TestMethod]
        public void ShouldReturnArray()
        {
            _sheet.Cells[1, 1].Value = "Brazil";
            _sheet.Cells[1, 2].Value = "Indonesia";
            _sheet.Cells[1, 3].Value = "Sweden";
            _sheet.Cells[2, 1].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[2, 3].Value = 3;
            _sheet.Cells[3, 1].Value = 4;
            _sheet.Cells[3, 2].Value = 5;
            _sheet.Cells[3, 3].Value = 6;

            _sheet.Cells["D4"].Formula = "XLOOKUP(\"Sweden\",A1:C1,A2:C3)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells["D4"].Value);
            Assert.AreEqual(6, _sheet.Cells["D5"].Value);
        }

        [DataTestMethod]
        [DataRow("*A*", "1", 1)]
        [DataRow("*A*", "3", -1)]
        public void ReverseSearchTestWildcard(string lookupValue, string expected, int searchMode)
        {
            _sheet.Cells[1, 1].Value = "ABC";
            _sheet.Cells[1, 2].Value = "1";
            _sheet.Cells[2, 1].Value = "DDD";
            _sheet.Cells[2, 2].Value = "2";
            _sheet.Cells[3, 1].Value = "ABC";
            _sheet.Cells[3, 2].Value = "3";

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(\"{lookupValue}\",A1:A3,B1:B3, \"Not found\", 2, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [DataTestMethod]
        [DataRow(0d, 1d, 0, 1)]
        [DataRow(11d, 2d, -1, 1)]
        [DataRow(11d, 3d, 1, 1)]
        public void HorizontalNumeric(double lookupValue, double expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 0;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 20;
            _sheet.Cells[2, 1].Value = 1d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 3d;

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP({lookupValue},A1:C1,A2:C2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value);
        }

        [DataTestMethod, Ignore]
        [DataRow(11d, 2d, -1, -2)]
        public void BinarySearchAsc(double lookupValue, double expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 0;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 20;
            _sheet.Cells[2, 1].Value = 1d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 3d;

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP({lookupValue},A1:C1,A2:C2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value);
        }

        [DataTestMethod, Ignore]
        [DataRow(0d, 1d, 0, -2)]
        [DataRow(11d, 3d, 1, -2)]
        public void BinarySearchDesc(double lookupValue, double expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 20;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 0;
            _sheet.Cells[2, 1].Value = 3d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 1d;

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP({lookupValue},A1:C1,A2:C2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, (double)_sheet.Cells["F2"].Value);
        }

        [TestMethod]
        public void TestBinarySearchUtilDesc()
        {
            var comparer = new XlookupObjectComparer(XLookupMatchMode.ExactMatch);
            var list = new List<XlookupSearchItem>
            {
                new XlookupSearchItem(20, 0),
                new XlookupSearchItem(17, 2),
                new XlookupSearchItem(15, 1),
                new XlookupSearchItem(14, 2),
                new XlookupSearchItem(13, 2),
                new XlookupSearchItem(12, 2),
                new XlookupSearchItem(10, 2),
                new XlookupSearchItem(9, 2),
                new XlookupSearchItem(8, 2),
                new XlookupSearchItem(5, 3),
                new XlookupSearchItem(1, 4)
            };
            var ix = XLookupBinarySearch.SearchDesc(10, list.ToArray(), comparer);
            Assert.AreEqual(6, ix);
            ix = XLookupBinarySearch.SearchDesc(20, list.ToArray(), comparer);
            Assert.AreEqual(0, ix);
            ix = XLookupBinarySearch.SearchDesc(1, list.ToArray(), comparer);
            Assert.AreEqual(10, ix);
            ix = XLookupBinarySearch.SearchDesc(5, list.ToArray(), comparer);
            Assert.AreEqual(9, ix);
            ix = XLookupBinarySearch.SearchDesc(0, list.ToArray(), comparer);
            Assert.AreEqual(-12, ix);
            ix = XLookupBinarySearch.SearchDesc(7, list.ToArray(), comparer);
            Assert.AreEqual(-10, ix);
            ix = XLookupBinarySearch.SearchDesc(21, list.ToArray(), comparer);
            Assert.AreEqual(-1, ix);
        }

        [TestMethod]
        public void TestBinarySearchUtilAsc()
        {
            var comparer = new XlookupObjectComparer(XLookupMatchMode.ExactMatch);
            var list = new List<XlookupSearchItem>
            {
                new XlookupSearchItem(1, 0),
                new XlookupSearchItem(5, 2),
                new XlookupSearchItem(8, 1),
                new XlookupSearchItem(9, 2),
                new XlookupSearchItem(10, 2),
                new XlookupSearchItem(12, 2),
                new XlookupSearchItem(13, 2),
                new XlookupSearchItem(14, 2),
                new XlookupSearchItem(15, 2),
                new XlookupSearchItem(17, 3),
                new XlookupSearchItem(20, 4)
            };
            var arr = list.ToArray();
            var ix = XLookupBinarySearch.Search(10, arr, comparer);
            Assert.AreEqual(4, ix);
            ix = XLookupBinarySearch.Search(20, arr, comparer);
            Assert.AreEqual(10, ix);
            ix = XLookupBinarySearch.Search(1, arr, comparer);
            Assert.AreEqual(0, ix);
            ix = XLookupBinarySearch.Search(5, arr, comparer);
            Assert.AreEqual(1, ix);
            ix = XLookupBinarySearch.Search(7, list.ToArray(), comparer);
            Assert.AreEqual(-3, ix);
        }
    }
}
