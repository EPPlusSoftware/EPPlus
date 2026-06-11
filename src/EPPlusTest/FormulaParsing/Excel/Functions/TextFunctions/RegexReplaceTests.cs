using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    /// <summary>
    /// Tests for the REGEXREPLACE function. All expected values are verified against
    /// the calculation of Excel desktop (see REGEX verification workbook).
    /// </summary>
    [TestClass]
    public class RegexReplaceTests : TestBase
    {
        // -------------------------------------------------------------------
        // occurrence argument
        // 0 or omitted = replace all, positive N = the Nth match,
        // negative N = the Nth match counted from the end.
        // Out-of-range occurrences leave the text unchanged.
        // -------------------------------------------------------------------

        [TestMethod]
        public void Occurrence_ZeroReplacesAll()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b-c-d";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",0)";
                sheet.Calculate();

                Assert.AreEqual("a+b+c+d", sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void Occurrence_OmittedReplacesAll()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b-c-d";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"-\",\"+\")";
                sheet.Calculate();

                Assert.AreEqual("a+b+c+d", sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void Occurrence_PositiveReplacesNthMatch()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b-c-d";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",1)";
                sheet.Cells["B2"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",2)";
                sheet.Cells["B3"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",3)";
                sheet.Calculate();

                Assert.AreEqual("a+b-c-d", sheet.Cells["B1"].Value);
                Assert.AreEqual("a-b+c-d", sheet.Cells["B2"].Value);
                Assert.AreEqual("a-b-c+d", sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void Occurrence_NegativeCountsFromEnd()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b-c-d";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",-1)";
                sheet.Cells["B2"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",-2)";
                sheet.Calculate();

                Assert.AreEqual("a-b-c+d", sheet.Cells["B1"].Value); // last match
                Assert.AreEqual("a-b+c-d", sheet.Cells["B2"].Value); // second from end
            }
        }

        [TestMethod]
        public void Occurrence_OutOfRangeLeavesTextUnchanged()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b-c-d";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",99)";
                sheet.Cells["B2"].Formula = "REGEXREPLACE(A1,\"-\",\"+\",-99)";
                sheet.Calculate();

                Assert.AreEqual("a-b-c-d", sheet.Cells["B1"].Value);
                Assert.AreEqual("a-b-c-d", sheet.Cells["B2"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Back references
        // -------------------------------------------------------------------

        [TestMethod]
        public void ValidBackReference_IsResolved()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "2026-Q2";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"(\\d{4})-(\\w+)\",\"$2_$1\")";
                sheet.Calculate();

                Assert.AreEqual("Q2_2026", sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void InvalidBackReference_ReturnsValueError()
        {
            // $1 without any capturing group, and $2 when only one group exists.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "2026";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"[0-9]+\",\"$1\")";
                sheet.Cells["B2"].Formula = "REGEXREPLACE(A1,\"([0-9]+)\",\"$2\")";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B2"].Value);
            }
        }

        [TestMethod]
        public void LiteralDollarSign_ViaDoubleDollar()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "price: $5";

                // "$$" -> literal $, "$1" -> group 1 ("5"), so the result is unchanged "price: $5".
                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"\\$(\\d)\",\"$$$1\")";
                sheet.Calculate();

                Assert.AreEqual("price: $5", sheet.Cells["B1"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Pattern semantics
        // -------------------------------------------------------------------

        [TestMethod]
        public void UnescapedDot_MatchesEveryCharacter()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a.b.c";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\".\",\"-\")";
                sheet.Calculate();

                Assert.AreEqual("-----", sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void EmptyPattern_InsertsReplacementAtEveryPosition()
        {
            // Excel treats an empty pattern as an empty match at every position and
            // inserts the replacement: "abc" with "x" -> "xaxbxcx".
            // KNOWN DIVERGENCE: the current scalar branch has a guard
            // (text != null && pattern == string.Empty) that returns #VALUE! instead.
            // This test is expected to FAIL until the empty-pattern guard is removed
            // (pending team decision on whether to match Excel parity).
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "abc";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"\",\"x\")";
                sheet.Calculate();

                Assert.AreEqual("xaxbxcx", sheet.Cells["B1"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Invalid arguments
        // -------------------------------------------------------------------

        [TestMethod]
        public void InvalidPattern_ReturnsValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"[\",\"x\")";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void InvalidCaseArgument_ReturnsValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                // signature: REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])
                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"k\",\"x\",0,2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void Range_ReplacementOutOfRange_ReturnsNA()
        {
            // When the replacement range is shorter than the text range, the unmatched row
            // gets #N/A (like REGEXTEST/REGEXEXTRACT). Current code computes an empty
            // replacement instead, producing "ef". Expected to FAIL until the range fix lands.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a-b";
                sheet.Cells["A2"].Value = "c-d";
                sheet.Cells["A3"].Value = "e-f";
                sheet.Cells["C1"].Value = "-";
                sheet.Cells["D1"].Value = "+";
                sheet.Cells["D2"].Value = "*";

                sheet.Cells["E1"].Formula = "REGEXREPLACE(A1:A3,C1,D1:D2)";
                sheet.Calculate();

                Assert.AreEqual("a+b", sheet.Cells["E1"].Value);
                Assert.AreEqual("c*d", sheet.Cells["E2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E3"].Value);
            }
        }

        [TestMethod]
        public void Range_InvalidPattern_FailsPerCell()
        {
            // Excel isolates the invalid pattern to its own cell (#VALUE!) and computes
            // the rest. Current code lets the exception bubble, failing the whole array.
            // The last row (after the error cell) verifies the loop does not break early.
            // Expected to FAIL until per-cell error handling lands.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linkoping";
                sheet.Cells["A3"].Value = "Orebro";
                sheet.Cells["A4"].Value = "Malmo";
                sheet.Cells["B1"].Value = "S";
                sheet.Cells["B2"].Value = "[A-Z]";
                sheet.Cells["B3"].Value = "[";   // invalid pattern
                sheet.Cells["B4"].Value = "M";

                sheet.Cells["E1"].Formula = "REGEXREPLACE(A1:A4,B1:B4,\"x\")";
                sheet.Calculate();

                Assert.AreEqual("xtockholm", sheet.Cells["E1"].Value);
                Assert.AreEqual("xinkoping", sheet.Cells["E2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E3"].Value);
                Assert.AreEqual("xalmo", sheet.Cells["E4"].Value); // computed AFTER the error cell
            }
        }
    }
}