using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class HlookupTests
    {
        [TestMethod]
        public void HLookupShouldReturnResultFromMatchingRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(2,A1:B2,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(5, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnResultFromMatchingRow_Array()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1:G1"].CreateArrayFormula("HLOOKUP(A1:B1,A1:B2,2)");
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
                Assert.AreEqual(5, sheet.Cells["G1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnResultFromMatchingRow_Wildcard()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(\"*B*\",A1:B2,2,0)";
                sheet.Cells[1, 1].Value = "ABC";
                sheet.Cells[1, 2].Value = "DEF";
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(2,A1:B2,2,false)";
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();
                var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
                Assert.AreEqual(expectedResult, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "HLOOKUP(1,A1:B2,2,true)";
                sheet.Cells[1, 1].Value = 2;
                sheet.Cells[1, 2].Value = 3;
                sheet.Cells[2, 1].Value = 3;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();
                var naError = ExcelErrorValue.Create(eErrorType.NA);
                Assert.AreEqual(naError, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void HLookupApproximateMatchShouldHandleBlankColumnsAroundData()
        {
            // Regression test: approximate match (range_lookup = TRUE) for HLOOKUP where the
            // key row is preceded by blank columns. Mirrors the VLOOKUP case but horizontally.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            // Blank columns A-C, keys in D1:H1, results in D2:H2.
            ws.Cells["D1"].Value = 10;
            ws.Cells["E1"].Value = 20;
            ws.Cells["F1"].Value = 30;
            ws.Cells["G1"].Value = 40;
            ws.Cells["H1"].Value = 50;
            ws.Cells["D2"].Value = 100;
            ws.Cells["E2"].Value = 200;
            ws.Cells["F2"].Value = 300;
            ws.Cells["G2"].Value = 400;
            ws.Cells["H2"].Value = 500;

            // Approximate match between keys returns the value of the largest key <= lookup.
            ws.Cells["A4"].Formula = "HLOOKUP(35,A1:H2,2,TRUE)";
            // Exact matches at the edges of the data block.
            ws.Cells["A5"].Formula = "HLOOKUP(10,A1:H2,2,TRUE)";
            ws.Cells["A6"].Formula = "HLOOKUP(50,A1:H2,2,TRUE)";
            // Lookup value smaller than every key returns #N/A.
            ws.Cells["A7"].Formula = "HLOOKUP(5,A1:H2,2,TRUE)";
            ws.Calculate();

            Assert.AreEqual(300, ws.Cells["A4"].Value);
            Assert.AreEqual(100, ws.Cells["A5"].Value);
            Assert.AreEqual(500, ws.Cells["A6"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["A7"].Value);
        }

        [TestMethod]
        public void HLookupApproximateMatch_LeadingBlankColumns()
        {
            // Sorted ascending data preceded by blank columns. The leading blanks must
            // be skipped so the search finds the data, matching Excel.
            //   C1=10, D1=20, E1=30, F1=40, G1=50  (A1, B1 blank), results in row 2.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            // A1, B1 intentionally left blank
            ws.Cells["C1"].Value = 10;
            ws.Cells["D1"].Value = 20;
            ws.Cells["E1"].Value = 30;
            ws.Cells["F1"].Value = 40;
            ws.Cells["G1"].Value = 50;

            ws.Cells["C2"].Value = 100;
            ws.Cells["D2"].Value = 200;
            ws.Cells["E2"].Value = 300;
            ws.Cells["F2"].Value = 400;
            ws.Cells["G2"].Value = 500;

            ws.Cells["A4"].Formula = "HLOOKUP(5,A1:G2,2,TRUE)";   // below the first value
            ws.Cells["A5"].Formula = "HLOOKUP(10,A1:G2,2,TRUE)";  // exact, first value
            ws.Cells["A6"].Formula = "HLOOKUP(15,A1:G2,2,TRUE)";  // approximate -> 10
            ws.Cells["A7"].Formula = "HLOOKUP(50,A1:G2,2,TRUE)";  // exact, last value
            ws.Cells["A8"].Formula = "HLOOKUP(55,A1:G2,2,TRUE)";  // above the last value -> 50
            ws.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["A4"].Value);
            Assert.AreEqual(100, ws.Cells["A5"].Value);
            Assert.AreEqual(100, ws.Cells["A6"].Value);
            Assert.AreEqual(500, ws.Cells["A7"].Value);
            Assert.AreEqual(500, ws.Cells["A8"].Value);
        }

        [TestMethod]
        public void HLookupApproximateMatch_InnerBlankBeforeLastValue()
        {
            // Sorted ascending data with an inner blank immediately before the last value:
            //   A1=10, B1=20, C1=<blank>, D1=30, E1=40, F1=<blank>, G1=50
            // HLOOKUP(50, A1:G2, 2, TRUE) must find the exact match 50 and return 500,
            // seeing past the inner blank at F1 instead of stopping there.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 10;
            ws.Cells["B1"].Value = 20;
            // C1 intentionally left blank
            ws.Cells["D1"].Value = 30;
            ws.Cells["E1"].Value = 40;
            // F1 intentionally left blank
            ws.Cells["G1"].Value = 50;

            ws.Cells["A2"].Value = 100;
            ws.Cells["B2"].Value = 200;
            ws.Cells["D2"].Value = 300;
            ws.Cells["E2"].Value = 400;
            ws.Cells["G2"].Value = 500;

            ws.Cells["A4"].Formula = "HLOOKUP(25,A1:G2,2,TRUE)";
            ws.Cells["A5"].Formula = "HLOOKUP(45,A1:G2,2,TRUE)";
            ws.Cells["A6"].Formula = "HLOOKUP(50,A1:G2,2,TRUE)";
            ws.Calculate();

            Assert.AreEqual(200, ws.Cells["A4"].Value);
            Assert.AreEqual(400, ws.Cells["A5"].Value);
            Assert.AreEqual(500, ws.Cells["A6"].Value);
        }

        [TestMethod]
        public void HLookupApproximateMatch_TrailingBlankColumns()
        {
            // Sorted ascending data followed by blank columns. The trailing blanks must
            // not affect the result, matching Excel.
            //   A1=10, B1=20, C1=30, D1=40, E1=50  (F1, G1 blank), results in row 2.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 10;
            ws.Cells["B1"].Value = 20;
            ws.Cells["C1"].Value = 30;
            ws.Cells["D1"].Value = 40;
            ws.Cells["E1"].Value = 50;
            // F1, G1 intentionally left blank

            ws.Cells["A2"].Value = 100;
            ws.Cells["B2"].Value = 200;
            ws.Cells["C2"].Value = 300;
            ws.Cells["D2"].Value = 400;
            ws.Cells["E2"].Value = 500;

            ws.Cells["A4"].Formula = "HLOOKUP(5,A1:G2,2,TRUE)";   // below the first value
            ws.Cells["A5"].Formula = "HLOOKUP(10,A1:G2,2,TRUE)";  // exact, first value
            ws.Cells["A6"].Formula = "HLOOKUP(35,A1:G2,2,TRUE)";  // approximate -> 30
            ws.Cells["A7"].Formula = "HLOOKUP(50,A1:G2,2,TRUE)";  // exact, last value
            ws.Cells["A8"].Formula = "HLOOKUP(55,A1:G2,2,TRUE)";  // above the last value -> 50
            ws.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["A4"].Value);
            Assert.AreEqual(100, ws.Cells["A5"].Value);
            Assert.AreEqual(300, ws.Cells["A6"].Value);
            Assert.AreEqual(500, ws.Cells["A7"].Value);
            Assert.AreEqual(500, ws.Cells["A8"].Value);
        }
    }
}
