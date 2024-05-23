using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class XLookupTests : TestBase
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
                _sheet.Cells["F2"].Formula = $"     (E2,A2:A11,C2:C11,\"{notFoundText}\")";
            }

            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value.ToString());
        }

        [DataTestMethod]
        [DataRow("Brazil", "+55", "Not found", 0)]
        [DataRow("Brasil", "Not found", "Not found", 0)]
        [DataRow("Bazil", "#N/A", null, 0)]
        [DataRow("Bazil", "+880", "Not found", -1)]
        [DataRow("Bsazil", "+86", "Not found", 1)]
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
                _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A1:A11,C1:C11,, {searchMode})";
            }
            else
            {
                _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A1:A11,C1:C11,\"{notFoundText}\", {searchMode})";
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
        public void ShouldReturnVerticalArray()
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

        [TestMethod]
        public void ShouldReturnHorizontalArray()
        {
            _sheet.Cells[1, 1].Value = "Brazil";
            _sheet.Cells[2, 1].Value = "Indonesia";
            _sheet.Cells[3, 1].Value = "Sweden";
            _sheet.Cells[1, 2].Value = 1;
            _sheet.Cells[2, 2].Value = 2;
            _sheet.Cells[3, 2].Value = 3;
            _sheet.Cells[1, 3].Value = 4;
            _sheet.Cells[2, 3].Value = 5;
            _sheet.Cells[3, 3].Value = 6;

            _sheet.Cells["D4"].Formula = "XLOOKUP(\"Sweden\",A1:A3,B1:C3)";
            _sheet.Calculate();

            Assert.AreEqual(3, _sheet.Cells["D4"].Value);
            Assert.AreEqual(6, _sheet.Cells["E4"].Value);
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
        [DataRow("*A*", "1", 1)]
        [DataRow("*A*", "3", -1)]
        public void ReverseSearchTestWildcardLargeRange(string lookupValue, string expected, int searchMode)
        {
            _sheet.Cells[1, 1].Value = "ABC";
            _sheet.Cells[1, 2].Value = "1";
            _sheet.Cells[2, 1].Value = "DDD";
            _sheet.Cells[2, 2].Value = "2";
            _sheet.Cells[3, 1].Value = "ABC";
            _sheet.Cells[3, 2].Value = "3";

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(\"{lookupValue}\",A1:A30000,B1:B30000, \"Not found\", 2, {searchMode})";


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

        [DataTestMethod]
        [DataRow(0d, 1d, 0, 1)]
        [DataRow(11d, 2d, -1, 1)]
        [DataRow(11d, 3d, 1, 1)]
        public void HorizontalNumericLarge(double lookupValue, double expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 0;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 20;
            _sheet.Cells[2, 1].Value = 1d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 3d;

            _sheet.Cells["E3"].Value = lookupValue;
            _sheet.Cells["F3"].Formula = $"XLOOKUP({lookupValue},A1:XFC1,A2:XFC2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F3"].Value);
        }

        [DataTestMethod]
        [DataRow(11d, 2d, -1, 2)]
        [DataRow(21d, 3d, -1, 2)]
        [DataRow(0d, "Not found", -1, 2)]
        public void BinarySearchAscNextSmaller(double lookupValue, object expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 20;
            _sheet.Cells[2, 1].Value = 1d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 3d;

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A1:C1,A2:C2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value);
        }

        [DataTestMethod]
        [DataRow(11d, 3d, 1, 2)]
        [DataRow(21d, "Not found", 1, 2)]
        [DataRow(0d, 1d, 1, 2)]
        public void BinarySearchAscNextLarger(double lookupValue, object expected, int matchMode, int searchMode)
        {
            _sheet.Cells[1, 1].Value = 1;
            _sheet.Cells[1, 2].Value = 10;
            _sheet.Cells[1, 3].Value = 20;
            _sheet.Cells[2, 1].Value = 1d;
            _sheet.Cells[2, 2].Value = 2d;
            _sheet.Cells[2, 3].Value = 3d;

            _sheet.Cells["E2"].Value = lookupValue;
            _sheet.Cells["F2"].Formula = $"XLOOKUP(E2,A1:C1,A2:C2, \"Not found\", {matchMode}, {searchMode})";


            _sheet.Calculate();

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value);
        }

        [DataTestMethod]
        [DataRow(0d, 1d, 0, -2)]
        [DataRow(11d, 3d, 1, -2)]
        [DataRow(21d, "Not found", 1, -2)]
        public void BinarySearchDesc(double lookupValue, object expected, int matchMode, int searchMode)
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

            Assert.AreEqual(expected, _sheet.Cells["F2"].Value);
        }

        [TestMethod]
        public void TestBinarySearchUtilDesc()
        {
            var comparer = new LookupComparer(LookupMatchMode.ExactMatch);
            var list = new List<LookupSearchItem>
            {
                new LookupSearchItem(20, 0),
                new LookupSearchItem(17, 2),
                new LookupSearchItem(15, 1),
                new LookupSearchItem(14, 2),
                new LookupSearchItem(13, 2),
                new LookupSearchItem(12, 2),
                new LookupSearchItem(10, 2),
                new LookupSearchItem(9, 2),
                new LookupSearchItem(8, 2),
                new LookupSearchItem(5, 3),
                new LookupSearchItem(1, 4)
            };
            var ix = LookupBinarySearch.SearchDesc(10, list.ToArray(), comparer);
            Assert.AreEqual(6, ix);
            ix = LookupBinarySearch.SearchDesc(20, list.ToArray(), comparer);
            Assert.AreEqual(0, ix);
            ix = LookupBinarySearch.SearchDesc(1, list.ToArray(), comparer);
            Assert.AreEqual(10, ix);
            ix = LookupBinarySearch.SearchDesc(5, list.ToArray(), comparer);
            Assert.AreEqual(9, ix);
            ix = LookupBinarySearch.SearchDesc(0, list.ToArray(), comparer);
            Assert.AreEqual(-12, ix);
            ix = LookupBinarySearch.SearchDesc(7, list.ToArray(), comparer);
            Assert.AreEqual(-10, ix);
            ix = LookupBinarySearch.SearchDesc(21, list.ToArray(), comparer);
            Assert.AreEqual(-1, ix);
        }

        [TestMethod]
        public void TestBinarySearchUtilAsc()
        {
            var comparer = new LookupComparer(LookupMatchMode.ExactMatch);
            var list = new List<LookupSearchItem>
            {
                new LookupSearchItem(1, 0),
                new LookupSearchItem(5, 2),
                new LookupSearchItem(8, 1),
                new LookupSearchItem(9, 2),
                new LookupSearchItem(10, 2),
                new LookupSearchItem(12, 2),
                new LookupSearchItem(13, 2),
                new LookupSearchItem(14, 2),
                new LookupSearchItem(15, 2),
                new LookupSearchItem(17, 3),
                new LookupSearchItem(20, 4)
            };
            var arr = list.ToArray();
            var ix = LookupBinarySearch.SearchAsc(10, arr, comparer);
            Assert.AreEqual(4, ix);
            ix = LookupBinarySearch.SearchAsc(20, arr, comparer);
            Assert.AreEqual(10, ix);
            ix = LookupBinarySearch.SearchAsc(1, arr, comparer);
            Assert.AreEqual(0, ix);
            ix = LookupBinarySearch.SearchAsc(5, arr, comparer);
            Assert.AreEqual(1, ix);
            ix = LookupBinarySearch.SearchAsc(7, list.ToArray(), comparer);
            Assert.AreEqual(-3, ix);
        }
        [TestMethod]
		public void XlookupSharedAndArray()
		{
			_sheet.Cells[1, 1].Value = "A";
			_sheet.Cells[2, 1].Value = "B";
			_sheet.Cells[3, 1].Value = "C";
			_sheet.Cells[1, 2].Value = "1";
			_sheet.Cells[2, 2].Value = "2";
			_sheet.Cells[3, 2].Value = "3";
			_sheet.Cells[1, 3].Value = "12";
			_sheet.Cells[2, 3].Value = "23";
			_sheet.Cells[3, 3].Value = "34";

			//_dateWs1.Cells["A5:A7"].SetFormula($"XLOOKUP(A1,$A$1:$A$3,$B$1:$C$3)", false);
			_sheet.Cells["A5"].Formula = $"XLOOKUP(A1,$A$1:$A$3,$B$1:$C$3)";
			_sheet.Cells["A6"].Formula = $"XLOOKUP(A2,$A$1:$A$3,$B$1:$C$3)";
			_sheet.Cells["A7"].Formula = $"XLOOKUP(A3,$A$1:$A$3,$B$1:$C$3)";
			_sheet.Cells["A15:A17"].Formula = $"= RANDARRAY(1,2)";

			_sheet.Calculate();
            SaveWorkbook("XLOOKUP.XLSX", _sheet._package);
			Assert.AreEqual("1", _sheet.Cells["A5"].Value);
			Assert.AreEqual("12", _sheet.Cells["B5"].Value);
			Assert.AreEqual("2", _sheet.Cells["A6"].Value);
			Assert.AreEqual("23", _sheet.Cells["B6"].Value);
			Assert.AreEqual("3", _sheet.Cells["A7"].Value);
			Assert.AreEqual("34", _sheet.Cells["B7"].Value);
		}

	}
}
