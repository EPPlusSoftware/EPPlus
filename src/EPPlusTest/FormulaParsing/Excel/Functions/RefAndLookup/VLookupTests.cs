using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class VLookupTests : TestBase
    {
        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(2,A1:B2,2)";
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(5, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow_Array()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1:F2"].CreateArrayFormula("VLOOKUP(A1:A2,A1:B2,2)");
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
                Assert.AreEqual(5, sheet.Cells["F2"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow_Wildcard()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(\"*B*\",A1:B2,2,0)";
                sheet.Cells[1, 1].Value = "ABC";
                sheet.Cells[1, 2].Value = 2;
                sheet.Cells[2, 1].Value = "DEF";
                sheet.Cells[2, 2].Value = 5;
                sheet.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(4,A1:B2,2,true)";
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = 5;
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(\"B\",A1:B2,2,true)";
                sheet.Cells[1, 1].Value = "A";
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = "C";
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void VLookupShouldIgnoreCase()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["F1"].Formula = "VLOOKUP(\"b\",A1:B2,2,true)";
                sheet.Cells[1, 1].Value = "A";
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 1].Value = "C";
                sheet.Cells[2, 2].Value = 4;
                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }
        [TestMethod]
        public void VLookupHeaderIncluded()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("sheet1");

                ws.Cells["A1"].Value = "Header";
                ws.Cells["A2"].Value = 1;
                ws.Cells["A3"].Value = 2;
                ws.Cells["A4"].Value = 3;
                ws.Cells["A5"].Value = 4;
                ws.Cells["B1"].Value = "Result";
                ws.Cells["B2"].Value = "Found1";
                ws.Cells["B3"].Value = "Found2";
                ws.Cells["B4"].Value = "Found3";
                ws.Cells["B5"].Value = "Found4";

                var result = ws.Calculate("VLOOKUP(1,A1:B5,2,TRUE)");
                Assert.AreEqual("Found1", result);
            }
        }

        [TestMethod]
        public void VlookupShouldHandleWholeColumn()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D1"].Value = 1;
                sheet.Cells["D2"].Value = 2;
                sheet.Cells["D3"].Value = 2;
                sheet.Cells["D4"].Value = 3;
                sheet.Cells["D5"].Value = 3;
                sheet.Cells["D6"].Value = 4;
                sheet.Cells["D7"].Value = 4;
                sheet.Cells["D8"].Value = 5;
                sheet.Cells["D9"].Value = 5;

                sheet.Cells["E1"].Value = "a";
                sheet.Cells["E2"].Value = "b";
                sheet.Cells["E3"].Value = "c";
                sheet.Cells["E4"].Value = "d";
                sheet.Cells["E5"].Value = "e";
                sheet.Cells["E6"].Value = "f";
                sheet.Cells["E7"].Value = "g";
                sheet.Cells["E8"].Value = "h";
                sheet.Cells["E9"].Value = "i";

                sheet.Cells["C10"].Formula = "VLOOKUP(3,D:E,2,FALSE)";
                sheet.Calculate();
                Assert.AreEqual("d", sheet.Cells["C10"].Value);
            }
        }


        [TestMethod]
        public void ApproximateShouldFindDate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["C1"].Formula = "TODAY()";

                sheet.Cells["A1"].Formula = "C1";
                sheet.Cells["A2"].Formula = "C1+1";
                sheet.Cells["A3"].Formula = "C1+3";
                sheet.Cells["A4"].Formula = "C1+7";

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = "VLOOKUP(C1,A1:B4,2,TRUE)";
                sheet.Calculate();
                Assert.AreEqual("a", sheet.Cells["D1"].Value);
            }
        }

        [DataTestMethod]
        [DataRow(1, "a")]
        [DataRow(5, "d")]
        public void ApproximateShouldFind(int find, string expected)
        {
            using (var package = OpenPackage("VlookupApprox_Finds.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = $"VLOOKUP({find},A1:B4,2,TRUE)";
                sheet.Calculate();

                Assert.AreEqual(expected, sheet.Cells["D1"].Value);
                //SaveAndCleanup(package);
            }
        }

        [TestMethod]

        public void ExactShouldNA()
        {
            using (var package = OpenPackage("VlookupExact_NotFound.xlsx",true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = $"VLOOKUP(5,A1:B4,2,FALSE)";
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.NAError, sheet.Cells["D1"].Value);
                //SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void ApproximateOutOfRangePositiveShouldRefError()
        {
            using (var package = OpenPackage("VlookupApprox_OutOfRangePositive_ReturnsRefError.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 4;

                sheet.Cells["C1"].Value = "a";
                sheet.Cells["C2"].Value = "b";
                sheet.Cells["C3"].Value = "c";
                sheet.Cells["C4"].Value = "d";

                sheet.Cells["D1"].Value = "aa";
                sheet.Cells["D2"].Value = "bb";
                sheet.Cells["D3"].Value = "cc";
                sheet.Cells["D4"].Value = "dd";

                sheet.Cells["E1"].Formula = $"VLOOKUP(2,B1:C4,{3},TRUE)"; // positive offset is out of range
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.RefError, sheet.Cells["E1"].Value);

                //SaveAndCleanup(package);
            }
        }

        [TestMethod]
        [DataRow(0)]
        [DataRow(-1)]
        public void ApproximateOutOfRangeNonPositiveShouldValueError(int offset)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 10;
                sheet.Cells["A2"].Value = 20;
                sheet.Cells["A3"].Value = 30;
                sheet.Cells["A4"].Value = 40;

                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 4;

                sheet.Cells["C1"].Value = "a";
                sheet.Cells["C2"].Value = "b";
                sheet.Cells["C3"].Value = "c";
                sheet.Cells["C4"].Value = "d";

                sheet.Cells["E1"].Formula = $"VLOOKUP(2,B1:C4,{offset},TRUE)";
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["E1"].Value);
                //SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ExactStringsShouldFind()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NewWs");

                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "b";
                sheet.Cells["A3"].Value = "c";
                sheet.Cells["A4"].Value = "d";

                sheet.Cells["B1"].Value = "aa";
                sheet.Cells["B2"].Value = "bb";
                sheet.Cells["B3"].Value = "cc";
                sheet.Cells["B4"].Value = "dd";

                sheet.Cells["C1"].Formula = $"VLOOKUP(\"c\", A1:B4, 2, FALSE)";
                sheet.Cells["C2"].Formula = $"VLOOKUP(\"d\", A1:B4, 2, FALSE)";

                sheet.Calculate();

                Assert.AreEqual("cc", sheet.Cells["C1"].Value);
                Assert.AreEqual("dd", sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void ApproximateStringsShouldFind()
        {
            using (var package = OpenPackage("VLOOKUP_approxStrings.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("NewWs");

                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "b";
                sheet.Cells["A3"].Value = "c";
                sheet.Cells["A4"].Value = "d";

                sheet.Cells["B1"].Value = "aa";
                sheet.Cells["B2"].Value = "bb";
                sheet.Cells["B3"].Value = "cc";
                sheet.Cells["B4"].Value = "dd";

                //"easy" to find
                sheet.Cells["C1"].Formula = $"VLOOKUP(\"ca\", A1:B4, 2, TRUE)";
                //Slightly harder to find
                sheet.Cells["C2"].Formula = $"VLOOKUP(\"da\", A1:B4, 2, TRUE)";

                sheet.Calculate();
                Assert.AreEqual("cc", sheet.Cells["C1"].Value);
                Assert.AreEqual("dd", sheet.Cells["C2"].Value);

                SaveAndCleanup(package);
            }
        }

        //Potentially support?
        //[TestMethod]
        //public void ApproximateMixedTypesByDateNumberFormat()
        //{
        //    using (var package = OpenPackage("VlookupApprox_MixedTypesByDateNumberFormat.xlsx",true))
        //    {
        //        // STAGING 
        //        var sheet = package.Workbook.Worksheets.Add("test");
        //        // mimicking error scenario with date value to reference
        //        sheet.Cells["A1"].Formula = "TODAY()";

        //        // VLOOKUP INPUT
        //        sheet.Cells["F1"].Formula = "A1+1";
        //        sheet.Cells["F1"].Style.Numberformat.Format = "[$-409]mmmm\\ d\\,\\ yyyy;@";

        //        // RANGE
        //        // mimicking error scenario with very specific, mixed values and formats 
        //        sheet.Cells["C1"].Value = "Today"; // Vlookup returns #N/A with this literal string value in the range
        //        //sheet.Cells["C1"].Formula= "A1"; // Vlookup returns expected result with this Date value in the range
        //        sheet.Cells["C1"].Style.Numberformat.Format = "[$-409]mmm\\-yy;@";
        //        sheet.Cells["C2"].Formula = "A1+1";
        //        sheet.Cells["C2"].Style.Numberformat.Format = "mm-dd-yy";
        //        sheet.Cells["C3"].Formula = "A1+3";
        //        sheet.Cells["C3"].Style.Numberformat.Format = "mm-dd-yy";
        //        sheet.Cells["C4"].Formula = "A1+7";
        //        sheet.Cells["C4"].Style.Numberformat.Format = "mm-dd-yy";

        //        sheet.Cells["D1"].Value = ".01";
        //        sheet.Cells["D1"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D2"].Value = ".02";
        //        sheet.Cells["D2"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D3"].Value = ".03";
        //        sheet.Cells["D3"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D4"].Value = ".04";
        //        sheet.Cells["D4"].Style.Numberformat.Format = "0%";

        //        // VLOOKUP OUTPUT
        //        sheet.Cells["F3"].Formula = "VLOOKUP(F1,C1:D4,2,TRUE)";

        //        //var logfile = new FileInfo(@"c:\temp\logfile.txt");
        //        //package.Workbook.FormulaParserManager.AttachLogger(logfile);

        //        sheet.Calculate();

        //        //var range = sheet.Cells["C1:D4"];
        //        //var val = sheet.Cells["F1"].Value;

        //        Assert.AreEqual(".02", sheet.Cells["F3"].Value);

        //        SaveAndCleanup(package);

        //    }
        //}

        [TestMethod]
        public void SC870_EpplusOnly()
        {
            using (var p = OpenPackage("EpplusNullAboveAndBelow.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("VLookupTest");
                List<int> searchValues = new List<int> { 1, 2, 4, 7, 11, 16, 21, 27 };
                List<int> resultValues = new List<int> { 400, 365, 315, 280, 250, 215, 200, 170 };

                ws.Cells["B6:B13"].LoadFromCollection(searchValues);
                ws.Cells["C6:C13"].LoadFromCollection(resultValues);

                ws.Cells["A11"].Value = 1;

                //Testing that VLookup (or rather binary search lookup) can handle values of 'null' in a range above and below target.
                ws.Cells["F6"].Formula = "VLOOKUP(A11, B:C, 2, TRUE)";

                ws.Calculate();

                var outputValue = ws.Cells["F6"].Value;
                Assert.AreEqual(400, outputValue);

                //Ensure it works for each of the values
                for (int i = 1; i < searchValues.Count; i++)
                {
                    var formulaCell = ws.Cells[6 + i, 6];
                    formulaCell.Formula = $"VLOOKUP({searchValues[i]}, B:C, 2, TRUE)";
                    formulaCell.Calculate();
                    Assert.AreEqual(resultValues[i], formulaCell.Value);
                }

                //Save Workbook
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void SC870()
        {
            using (var package = OpenTemplatePackage("s870.xlsx"))
            {
                var wb = package.Workbook;
                var worksheet = package.Workbook.Worksheets[0];

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    sheet.Hidden = eWorkSheetHidden.Visible;
                }

                worksheet.Cells["F15"].Formula = "VLOOKUP(B11, Salgsfragt!B:C, 2, TRUE)";

                var sWs = package.Workbook.Worksheets.GetByName("Salgsfragt");
                sWs.Cells["B4"].Value = null;
                sWs.Cells["B2"].Value = null;

                worksheet.Cells["F15"].Calculate();

                var someVal = worksheet.Cells["F15"].Value;
                var errorText = worksheet.Cells["D8"].Text;

                var cellEuItemPrice = worksheet.Cells["C18"];
                var cellEuTransportPrice = worksheet.Cells["C19"];
                var cellEuTotal = worksheet.Cells["C20"];

                var cellDKItemPrice = worksheet.Cells["C24"];
                var cellDKTransportPrice = worksheet.Cells["C25"];
                var cellDKTotal = worksheet.Cells["C26"];

                worksheet.Calculate();
                decimal tolerance = 0.1M;

                Assert.AreEqual(301.01M, (decimal)cellEuItemPrice.GetCellValue<double>(), tolerance);
                Assert.AreEqual(53.62M, (decimal)cellEuTransportPrice.GetCellValue<double>(), tolerance);
                Assert.AreEqual(354.62M, (decimal)cellEuTotal.GetCellValue<double>(), tolerance);

                Assert.AreEqual(2245.50M, (decimal)cellDKItemPrice.GetCellValue<double>(), tolerance);
                Assert.AreEqual(400M, (decimal)cellDKTransportPrice.GetCellValue<double>(), tolerance);
                Assert.AreEqual(2645.50M, (decimal)cellDKTotal.GetCellValue<double>(), tolerance);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void PriorAddressExpressionWorksheetShouldBeCleared()
        {
            using (var pck = OpenPackage("vlookuptest.xlsx", true))
            {
                #region firstWorksheet
                using var firstWorksheet = pck.Workbook.Worksheets.Add("firstWorksheet");

                firstWorksheet.SetValue("A1", 4000);
                firstWorksheet.Names.Add("search", new ExcelRange(firstWorksheet, "A1"));

                firstWorksheet.SetValue("B53", 0); firstWorksheet.SetValue("C53", -1); firstWorksheet.SetValue("D53", -1);
                firstWorksheet.SetValue("B54", 3500); firstWorksheet.SetValue("C54", -1); firstWorksheet.SetValue("D54", 151);
                firstWorksheet.SetValue("B55", 4500); firstWorksheet.SetValue("C55", -1); firstWorksheet.SetValue("D55", -1);

                firstWorksheet.SetFormula(2, 1, "VLOOKUP(firstWorksheet!search,$B$53:$D$55,3,1)");

                pck.Workbook.Calculate();

                Assert.AreEqual(151, firstWorksheet.Cells["A2"].Value);
                #endregion

                #region secondWorksheet
                using var secondWorksheet = pck.Workbook.Worksheets.Add("secondWorksheet");

                secondWorksheet.SetValue("B53", 0); secondWorksheet.SetValue("C53", -1); secondWorksheet.SetValue("D53", -1);
                secondWorksheet.SetValue("B54", 3500); secondWorksheet.SetValue("C54", -1); secondWorksheet.SetValue("D54", 251);
                secondWorksheet.SetValue("B55", 4500); secondWorksheet.SetValue("C55", -1); secondWorksheet.SetValue("D55", -1);

                secondWorksheet.SetFormula(2, 1, "VLOOKUP(firstWorksheet!search,$B$53:$D$55,3,1)");

                secondWorksheet.Calculate();

                Assert.AreEqual(251, secondWorksheet.Cells["A2"].Value);
                SaveAndCleanup(pck);
                #endregion
            }
        }
        [TestMethod]
        public void VlookupMemoryRange()
        {
            using (var p = OpenPackage("MemRange.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Ws1");
                //ws.Cells["A1:A10"].Formula = "ROW()+1";
                //ws.Cells["B1:B10"].Formula = "ROW()";
                ws.Cells["C1"].Formula = "VLOOKUP(\"b\",TRANSPOSE({\"a\",\"b\",\"c\";1,2,3}),2)";
                ws.Calculate();

                Assert.AreEqual(2, ws.Cells["C1"].Value);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void FullCols()
        {
            using (var package = OpenTemplatePackage("ExternalVLookupFullCols.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                ws.Calculate();
                Assert.AreEqual(6d, ws.Cells["A1"].Value);
                Assert.AreEqual(7d, ws.Cells["A2"].Value);
                Assert.AreEqual(8d, ws.Cells["A3"].Value);
                Assert.AreEqual(10d, ws.Cells["A4"].Value);
                Assert.AreEqual(13d, ws.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void VLookupApproximateMatchShouldHandleBlankRowsAroundData()
        {
            // Regression test: approximate match (range_lookup = TRUE) over a whole-column
            // reference where the data is preceded and followed by blank rows.
            // The lookup column is not trimmed/compacted, so leading blanks must be skipped
            // while the binary search still partitions the range the same way Excel does.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            // Blank rows 1-3, data in rows 4-8, blank rows below.
            ws.Cells["B4"].Value = 10;
            ws.Cells["B5"].Value = 20;
            ws.Cells["B6"].Value = 30;
            ws.Cells["B7"].Value = 40;
            ws.Cells["B8"].Value = 50;
            ws.Cells["C4"].Value = 100;
            ws.Cells["C5"].Value = 200;
            ws.Cells["C6"].Value = 300;
            ws.Cells["C7"].Value = 400;
            ws.Cells["C8"].Value = 500;

            // Approximate match between keys returns the value of the largest key <= lookup.
            ws.Cells["E1"].Formula = "VLOOKUP(35,B:C,2,TRUE)";
            // Exact matches at the edges of the data block.
            ws.Cells["E2"].Formula = "VLOOKUP(10,B:C,2,TRUE)";
            ws.Cells["E3"].Formula = "VLOOKUP(50,B:C,2,TRUE)";
            // Lookup value smaller than every key returns #N/A.
            ws.Cells["E4"].Formula = "VLOOKUP(5,B:C,2,TRUE)";
            ws.Calculate();

            Assert.AreEqual(300, ws.Cells["E1"].Value);
            Assert.AreEqual(100, ws.Cells["E2"].Value);
            Assert.AreEqual(500, ws.Cells["E3"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["E4"].Value);
        }

        [TestMethod]
        public void VLookupApproximateMatch_InnerBlankBeforeLastValue()
        {
            // Sorted ascending data with an inner blank immediately before the last value:
            //   A1=10, A2=20, A3=<blank>, A4=30, A5=40, A6=<blank>, A7=50
            // VLOOKUP(50, A1:B7, 2, TRUE) must find the exact match 50 and return 500,
            // as Excel does. The approximate binary search sees past the inner blank at
            // A6 to the value at A7 instead of treating the blank as a stopping point.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 10;
            ws.Cells["A2"].Value = 20;
            // A3 intentionally left blank
            ws.Cells["A4"].Value = 30;
            ws.Cells["A5"].Value = 40;
            // A6 intentionally left blank
            ws.Cells["A7"].Value = 50;

            ws.Cells["B1"].Value = 100;
            ws.Cells["B2"].Value = 200;
            ws.Cells["B4"].Value = 300;
            ws.Cells["B5"].Value = 400;
            ws.Cells["B7"].Value = 500;

            ws.Cells["D1"].Formula = "VLOOKUP(25,A1:B7,2,TRUE)";
            ws.Cells["D2"].Formula = "VLOOKUP(45,A1:B7,2,TRUE)";
            ws.Cells["D3"].Formula = "VLOOKUP(50,A1:B7,2,TRUE)";
            ws.Calculate();

            // These already agree with Excel.
            Assert.AreEqual(200, ws.Cells["D1"].Value);
            Assert.AreEqual(400, ws.Cells["D2"].Value);
            Assert.AreEqual(500, ws.Cells["D3"].Value);
        }

        [TestMethod]
        public void VLookupApproximateMatch_LeadingBlankRows()
        {
            // Sorted ascending data preceded by blank rows (e.g. a whole-column
            // reference where the data starts further down). The leading blanks must
            // be skipped so the search finds the data, matching Excel.
            //   A1=<blank>, A2=<blank>, A3=10, A4=20, A5=30, A6=40, A7=50
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            // A1, A2 intentionally left blank
            ws.Cells["A3"].Value = 10;
            ws.Cells["A4"].Value = 20;
            ws.Cells["A5"].Value = 30;
            ws.Cells["A6"].Value = 40;
            ws.Cells["A7"].Value = 50;

            ws.Cells["B3"].Value = 100;
            ws.Cells["B4"].Value = 200;
            ws.Cells["B5"].Value = 300;
            ws.Cells["B6"].Value = 400;
            ws.Cells["B7"].Value = 500;

            ws.Cells["D1"].Formula = "VLOOKUP(5,A1:B7,2,TRUE)";   // below the first value
            ws.Cells["D2"].Formula = "VLOOKUP(10,A1:B7,2,TRUE)";  // exact, first value
            ws.Cells["D3"].Formula = "VLOOKUP(15,A1:B7,2,TRUE)";  // approximate -> 10
            ws.Cells["D4"].Formula = "VLOOKUP(50,A1:B7,2,TRUE)";  // exact, last value
            ws.Cells["D5"].Formula = "VLOOKUP(55,A1:B7,2,TRUE)";  // above the last value -> 50
            ws.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["D1"].Value);
            Assert.AreEqual(100, ws.Cells["D2"].Value);
            Assert.AreEqual(100, ws.Cells["D3"].Value);
            Assert.AreEqual(500, ws.Cells["D4"].Value);
            Assert.AreEqual(500, ws.Cells["D5"].Value);
        }

        [TestMethod]
        public void VLookupApproximateMatch_TrailingBlankRows()
        {
            // Sorted ascending data followed by blank rows. The trailing blanks must
            // not affect the result, matching Excel.
            //   A1=10, A2=20, A3=30, A4=40, A5=50, A6=<blank>, A7=<blank>
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 10;
            ws.Cells["A2"].Value = 20;
            ws.Cells["A3"].Value = 30;
            ws.Cells["A4"].Value = 40;
            ws.Cells["A5"].Value = 50;
            // A6, A7 intentionally left blank

            ws.Cells["B1"].Value = 100;
            ws.Cells["B2"].Value = 200;
            ws.Cells["B3"].Value = 300;
            ws.Cells["B4"].Value = 400;
            ws.Cells["B5"].Value = 500;

            ws.Cells["D1"].Formula = "VLOOKUP(5,A1:B7,2,TRUE)";   // below the first value
            ws.Cells["D2"].Formula = "VLOOKUP(10,A1:B7,2,TRUE)";  // exact, first value
            ws.Cells["D3"].Formula = "VLOOKUP(35,A1:B7,2,TRUE)";  // approximate -> 30
            ws.Cells["D4"].Formula = "VLOOKUP(50,A1:B7,2,TRUE)";  // exact, last value
            ws.Cells["D5"].Formula = "VLOOKUP(55,A1:B7,2,TRUE)";  // above the last value -> 50
            ws.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["D1"].Value);
            Assert.AreEqual(100, ws.Cells["D2"].Value);
            Assert.AreEqual(300, ws.Cells["D3"].Value);
            Assert.AreEqual(500, ws.Cells["D4"].Value);
            Assert.AreEqual(500, ws.Cells["D5"].Value);
        }

        [TestMethod]
        public void VLookupApproximateMatch_UnsortedWithLeadingBlanks_MatchesExcel()
        {
            // Approximate match runs as a plain binary search over the whole range with
            // blanks kept in their original positions, exactly like Excel. On unsorted
            // data the result is whatever that binary search lands on - it is not
            // "correct" in a sorted sense, but it matches Excel, which is the contract.
            //
            // Data (2 leading blanks, then unsorted keys):
            //   A1=<blank>, A2=<blank>, A3=30, A4=10, A5=50, A6=20, A7=40
            //
            // Verified against Excel:
            //   VLOOKUP(35) -> key 20 (B6)   - binary search lands here, not on 30
            //   VLOOKUP(30) -> key 20 (B6)   - the exact 30 is never visited by the search
            //   VLOOKUP(5)  -> #N/A
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            // A1, A2 intentionally left blank
            ws.Cells["A3"].Value = 30;
            ws.Cells["A4"].Value = 10;
            ws.Cells["A5"].Value = 50;
            ws.Cells["A6"].Value = 20;
            ws.Cells["A7"].Value = 40;

            ws.Cells["B3"].Value = 300;
            ws.Cells["B4"].Value = 100;
            ws.Cells["B5"].Value = 500;
            ws.Cells["B6"].Value = 200;
            ws.Cells["B7"].Value = 400;

            ws.Cells["D1"].Formula = "VLOOKUP(35,A1:B7,2,TRUE)";
            ws.Cells["D2"].Formula = "VLOOKUP(30,A1:B7,2,TRUE)";
            ws.Cells["D3"].Formula = "VLOOKUP(5,A1:B7,2,TRUE)";
            ws.Calculate();

            Assert.AreEqual(200, ws.Cells["D1"].Value);
            Assert.AreEqual(200, ws.Cells["D2"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ws.Cells["D3"].Value);
        }

        [TestMethod]
        public void VLookupApproximateMatch_FullColumn_ScanStopsAtDimension()
        {
            // Debugging aid for the inner blank-skipping scan bound.
            //
            // A:B is a whole-column reference (1,048,576 rows) but the data only occupies
            // rows 1-5. For an approximate match the first midpoints land deep in the empty
            // tail (e.g. row ~524288). Without the dimension bound, the skip-right scan
            // would walk forward through hundreds of thousands of empty cells looking for a
            // value. With the bound (scan limited to the last value position), each such
            // midpoint does a single read and the whole lookup completes in ~20 reads.
            //
            // To verify in the debugger: set a breakpoint inside the inner
            // 'while (probe < scanLimit && cellValue == null)' loop in
            // LookupBinarySearch.SearchAscFullRange. For the early midpoints you should see
            // 'mid' in the hundred-thousands while 'scanLimit' equals the last value offset
            // (4), so the loop body never executes.
            //
            // VLOOKUP(35, A:B, 2, TRUE) -> largest key <= 35 is 30 -> 300.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");

            ws.Cells["A1"].Value = 10;
            ws.Cells["A2"].Value = 20;
            ws.Cells["A3"].Value = 30;
            ws.Cells["A4"].Value = 40;
            ws.Cells["A5"].Value = 50;

            ws.Cells["B1"].Value = 100;
            ws.Cells["B2"].Value = 200;
            ws.Cells["B3"].Value = 300;
            ws.Cells["B4"].Value = 400;
            ws.Cells["B5"].Value = 500;

            ws.Cells["D1"].Formula = "VLOOKUP(35,A:B,2,TRUE)";
            ws.Calculate();

            Assert.AreEqual(300, ws.Cells["D1"].Value);
        }
    }
}
