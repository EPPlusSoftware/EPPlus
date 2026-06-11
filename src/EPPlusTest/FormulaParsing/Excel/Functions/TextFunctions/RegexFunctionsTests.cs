using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    /// <summary>
    /// Integration / range-level tests for the REGEX functions.
    ///
    /// Scalar behavior, case sensitivity and the per-function range cases now live in
    /// RegexTestTests, RegexExtractTests and RegexReplaceTests. This file keeps the
    /// broader multi-column / broadcast / integration tests that have no equivalent there.
    ///
    /// STATUS NOTES:
    ///  - Two tests are intentionally RED until the replacement-grammar fix lands and act as
    ///    regression guards: RegexReplaceValueError and Replacement_BackslashGroupReference_MatchesExcel.
    ///    Do not "fix" them by changing the assert - they encode verified Excel behavior.
    ///  - Tests marked "UNVERIFIED ASSERT" below were written from assumptions, not confirmed
    ///    against Excel desktop. Their expected values should be verified (and may shift once the
    ///    range / broadcast / replacement-grammar fixes land). Treat them as provisional.
    /// </summary>
    [TestClass]
    public class RegexFunctionsTests : TestBase
    {
        // NOTE: This test mixes one text/pattern pair per row. The expected values follow
        // standard regex semantics and are reasonably safe, but the row-by-row pairing is
        // hard to read - consider documenting each pair if this test is kept long-term.
        [TestMethod]
        public void RegexTest()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "^S";
                sheet.Cells["B3"].Value = "Q[0-9]";
                sheet.Cells["B4"].Value = "202[456]";
                sheet.Cells["B5"].Value = "^[0-9]{5}$";
                sheet.Cells["B6"].Value = "[A-ZÅÄÖ][a-zåäö]+";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A6, B1:B6)";
                sheet.Calculate();
                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(false, sheet.Cells["D2"].Value);
                Assert.AreEqual(false, sheet.Cells["D3"].Value);
                Assert.AreEqual(false, sheet.Cells["D4"].Value);
                Assert.AreEqual(false, sheet.Cells["D5"].Value);
                Assert.AreEqual(true, sheet.Cells["D6"].Value);
            }
        }

        [TestMethod]
        public void RegexTestMultiplCols()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "Linköping";
                sheet.Cells["B3"].Value = "Örebro";
                sheet.Cells["B4"].Value = "Stockholm";
                sheet.Cells["B5"].Value = "Örebro";
                sheet.Cells["B6"].Value = "Linköping";

                sheet.Cells["C1"].Value = "Stockholm";
                sheet.Cells["C2"].Value = "^S";
                sheet.Cells["C3"].Value = "Q[0-9]";
                sheet.Cells["C4"].Value = "202[456]";
                sheet.Cells["C5"].Value = "^[0-9]{5}$";
                sheet.Cells["C6"].Value = "[A-ZÅÄÖ][a-zåäö]+";

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:B6, C1:C6)";
                sheet.Calculate();
                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(false, sheet.Cells["D2"].Value);
                Assert.AreEqual(false, sheet.Cells["D3"].Value);
                Assert.AreEqual(false, sheet.Cells["D4"].Value);
                Assert.AreEqual(false, sheet.Cells["D5"].Value);
                Assert.AreEqual(true, sheet.Cells["D6"].Value);

                Assert.AreEqual(true, sheet.Cells["E1"].Value);
                Assert.AreEqual(false, sheet.Cells["E2"].Value);
                Assert.AreEqual(false, sheet.Cells["E3"].Value);
                Assert.AreEqual(false, sheet.Cells["E4"].Value);
                Assert.AreEqual(false, sheet.Cells["E5"].Value);
                Assert.AreEqual(true, sheet.Cells["E6"].Value);
            }
        }

        // UNVERIFIED ASSERT: the #N/A fill pattern for uneven ranges with numeric input was not
        // confirmed against Excel. This overlaps the broadcast behavior we are still fixing -
        // verify the expected values before trusting this as a regression guard.
        [TestMethod]
        public void RegexUnevenInputRanges()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";
                sheet.Cells["A4"].Value = "Stockholm";
                sheet.Cells["A5"].Value = "Örebro";
                sheet.Cells["A6"].Value = "Linköping";

                sheet.Cells["B1"].Value = "Stockholm";
                sheet.Cells["B2"].Value = "Linköping";
                sheet.Cells["B3"].Value = "Örebro";
                sheet.Cells["B4"].Value = "Stockholm";
                sheet.Cells["B5"].Value = "Örebro";
                sheet.Cells["B6"].Value = "Linköping";

                sheet.Cells["C1"].Value = 2026;
                sheet.Cells["C2"].Value = 2026;
                sheet.Cells["C3"].Value = 2025;
                sheet.Cells["C4"].Value = 2025;
                sheet.Cells["C5"].Value = 2025;
                sheet.Cells["C6"].Value = 2024;

                sheet.Cells["D4"].Value = "202[456]";
                sheet.Cells["D5"].Value = "^[0-9]{5}$";
                sheet.Cells["D6"].Value = "[A-ZÅÄÖ][a-zåäö]+";
                sheet.Cells["D7"].Value = "[0-9]+";

                sheet.Cells["E1"].Formula = "REGEXTEST(A1:C6, D4:D7)";
                sheet.Calculate();

                Assert.AreEqual(false, sheet.Cells["E1"].Value);
                Assert.AreEqual(false, sheet.Cells["E2"].Value);
                Assert.AreEqual(true, sheet.Cells["E3"].Value);
                Assert.AreEqual(false, sheet.Cells["E4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E6"].Value);

                Assert.AreEqual(false, sheet.Cells["F1"].Value);
                Assert.AreEqual(false, sheet.Cells["F2"].Value);
                Assert.AreEqual(true, sheet.Cells["F3"].Value);
                Assert.AreEqual(false, sheet.Cells["F4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F6"].Value);

                Assert.AreEqual(true, sheet.Cells["G1"].Value);
                Assert.AreEqual(false, sheet.Cells["G2"].Value);
                Assert.AreEqual(false, sheet.Cells["G3"].Value);
                Assert.AreEqual(true, sheet.Cells["G4"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["G6"].Value);
            }
        }

        // UNVERIFIED ASSERT: the email-pattern results over a range were not confirmed against
        // Excel. The pattern also has an unescaped dot before the TLD, so it matches more loosely
        // than it looks. Verify the expected values (and consider escaping the dot) before trusting.
        [TestMethod]
        public void RegexExtract()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Kossa Mail@mail.se";
                sheet.Cells["A2"].Value = "Får enmail@mef.se sd";
                sheet.Cells["A3"].Value = "mailens@hemma.com";
                sheet.Cells["A4"].Value = "mail@se.se";
                sheet.Cells["A5"].Value = "Tupp ska gala gmail@adress.net dwqdw";
                sheet.Cells["A6"].Value = "Katt";

                sheet.Cells["B1"].Value = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A6, B1)";
                sheet.Calculate();
                Assert.AreEqual("Mail@mail.se", sheet.Cells["D1"].Value);
                Assert.AreEqual("enmail@mef.se sd", sheet.Cells["D2"].Value);
                Assert.AreEqual("mailens@hemma.com", sheet.Cells["D3"].Value);
                Assert.AreEqual("mail@se.se", sheet.Cells["D4"].Value);
                Assert.AreEqual("gmail@adress.net dwqdw", sheet.Cells["D5"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D6"].Value);
            }
        }

        // Verifies that range mode 1 returns only the FIRST match per row (no horizontal spill
        // inside a range result). Uses weak AreNotEqual assertions - the stronger per-cell checks
        // live in RegexExtractTests.Range_Mode1_NoMatchCell_ReturnsNA.
        [TestMethod]
        public void RegexExtractShouldReturnSingleWithReturnMode1()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "Just #fitness finished 5k! #running";
                sheet.Cells["A2"].Value = "Look at this picture #nature #instagram";
                sheet.Cells["B1"].Value = "#\\w+";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2, B1, 1)";
                sheet.Calculate();
                Assert.AreEqual("#fitness", sheet.Cells["D1"].Value);
                Assert.AreNotEqual("#running", sheet.Cells["E1"].Value);
                Assert.AreEqual("#nature", sheet.Cells["D2"].Value);
                Assert.AreNotEqual("#instagram", sheet.Cells["E2"].Value);
            }
        }

        // Verifies that range mode 2 returns only the FIRST group per row.
        [TestMethod]
        public void RegexExtractShouldReturnSingleWithReturnMode2()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "9183-Green-M";
                sheet.Cells["A2"].Value = "2546-Black-XL";

                sheet.Cells["B1"].Value = "(\\d{4})-(\\w+)-(\\w+)";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2, B1, 2)";
                sheet.Calculate();

                Assert.AreEqual("9183", sheet.Cells["D1"].Value);
                Assert.AreEqual("2546", sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void RegexReplace()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "044-5654-6546";

                sheet.Cells["B1"].Value = "[^0-9]";

                sheet.Cells["D1"].Formula = "REGEXREPLACE(A1,B1,C1)";
                sheet.Calculate();

                Assert.AreEqual("04456546546", sheet.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void RegexReplaceWithOccurrance()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "044-5654-6546";

                sheet.Cells["B1"].Value = "[^0-9]";

                sheet.Cells["D1"].Formula = "REGEXREPLACE(A1,B1,C1, -1)";
                sheet.Calculate();

                Assert.AreEqual("044-56546546", sheet.Cells["D1"].Value);
            }
        }

        // UNVERIFIED ASSERT: this range case overlaps the broadcast / out-of-range #N/A fix we are
        // landing. The replacement range (B1:C2) reuses pattern-like text and the dimensions are
        // uneven - verify the expected values against Excel, as the fix may change row 3.
        [TestMethod]
        public void RegexReplaceRangeInput()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "044-5654-6546";
                sheet.Cells["A2"].Value = "0546-4654-565";

                sheet.Cells["D1"].Value = "[^0-9]";
                sheet.Cells["D2"].Value = "[^0-9]";

                sheet.Cells["E1"].Formula = "REGEXREPLACE(A1:B2,D1:D3, B1:C2)";
                sheet.Calculate();

                Assert.AreEqual("04456546546", sheet.Cells["E1"].Value);
                Assert.AreEqual("05464654565", sheet.Cells["E2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["E3"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["F3"].Value);
            }
        }

        // REGRESSION GUARD - intentionally RED until the replacement-grammar fix lands.
        // C1 is blank (empty pattern) and D1 holds a replacement containing \d \w, which Excel
        // rejects with #VALUE!. Current code inserts it literally. Do not change this assert.
        [TestMethod]
        public void RegexReplaceValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                sheet.Cells["A1"].Value = "044-5654-6546";
                sheet.Cells["D1"].Value = "(\\d{4})-(\\w+)-(\\w+)";
                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,C1,D1)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        // REGRESSION GUARD - intentionally RED until the replacement-grammar fix lands.
        // Verified against Excel: \1 is a back reference, so the result is "ax1y".
        // Current code passes the replacement straight to Regex.Replace and yields "ax\1y".
        [TestMethod]
        public void Replacement_BackslashGroupReference_MatchesExcel()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1";
                sheet.Cells["B1"].Formula = "REGEXREPLACE(A1,\"(\\d)\",\"x\\1y\")";
                sheet.Calculate();
                Assert.AreEqual("ax1y", sheet.Cells["B1"].Value);
            }
        }

        // UNVERIFIED ASSERT (partly): Test 2 below assumes a BLANK CELL pattern (C2/C3) behaves
        // like an empty-string pattern (insertion). The scalar empty-string case is verified, and
        // blank cells coerced to empty string also gave insertion in scalar tests, but the
        // range-mode blank-cell behavior here was not directly confirmed against Excel. The fix to
        // per-cell error handling and replacement grammar may also affect these rows - verify.
        [TestMethod]
        public void RegexReplaceInvalidBackreference()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");

                // Text column (A1:A3)
                sheet.Cells["A1"].Value = "2026-Stockholm-Q2";
                sheet.Cells["A2"].Value = "2025-Linkoping-Q1";
                sheet.Cells["A3"].Value = "2024-Orebro-Q4";

                // Pattern: only C1 set (3 groups), C2 and C3 are empty
                sheet.Cells["C1"].Value = @"(\d{4})-(\w+)-(\w+)";

                // Test 1: replacement with back references ($3_$1)
                // row 1 has groups -> ok, rows 2-3 empty pattern -> 0 groups -> #VALUE!
                sheet.Cells["E1"].Formula = "REGEXREPLACE(A1:A3, C1:C3, \"$3_$1\")";

                // Test 2: same setup, replacement WITHOUT back reference ("s")
                // empty pattern matches every position -> "s" inserted everywhere
                sheet.Cells["G1"].Formula = "REGEXREPLACE(A1:A3, C1:C3, \"s\")";

                // Test 3: scalar - valid pattern + valid back reference
                sheet.Cells["I1"].Formula = "REGEXREPLACE(A1, C1, \"$3_$1\")";

                // Test 4: scalar - pattern WITHOUT groups + back reference $1 -> #VALUE!
                sheet.Cells["I2"].Formula = "REGEXREPLACE(A1, \"[0-9]+\", \"$1\")";
                sheet.Calculate();

                // Test 1
                Assert.AreEqual("Q2_2026", sheet.Cells["E1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E3"].Value);

                // Test 2
                Assert.AreEqual("s", sheet.Cells["G1"].Value);
                Assert.AreEqual("s2s0s2s5s-sLsisnsksospsisnsgs-sQs1s", sheet.Cells["G2"].Value);
                Assert.AreEqual("s2s0s2s4s-sOsrsesbsrsos-sQs4s", sheet.Cells["G3"].Value);

                // Test 3
                Assert.AreEqual("Q2_2026", sheet.Cells["I1"].Value);

                // Test 4: $1 does not exist ([0-9]+ has no groups) -> #VALUE!
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["I2"].Value);
            }
        }
    }
}