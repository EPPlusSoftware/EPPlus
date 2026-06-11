using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    /// <summary>
    /// Tests for the REGEXTEST function. All expected values are verified against
    /// the calculation of Excel desktop (see REGEX verification workbook).
    /// </summary>
    [TestClass]
    public class RegexTestTests : TestBase
    {
        // -------------------------------------------------------------------
        // Case sensitivity
        // Verified in Excel: argument 0 = case-SENSITIVE, argument 1 = case-INSENSITIVE.
        // (This matches (RegexOptions)0 = None and (RegexOptions)1 = IgnoreCase.)
        // -------------------------------------------------------------------

        [TestMethod]
        public void CaseSensitivity_Arg0_IsCaseSensitive()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"stockholm\",0)";
                sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"STOCKHOLM\",0)";
                sheet.Cells["B3"].Formula = "REGEXTEST(A1,\"Stockholm\",0)";
                sheet.Calculate();

                Assert.AreEqual(false, sheet.Cells["B1"].Value); // different case does not match
                Assert.AreEqual(false, sheet.Cells["B2"].Value);
                Assert.AreEqual(true, sheet.Cells["B3"].Value);  // exact case matches
            }
        }

        [TestMethod]
        public void CaseSensitivity_Arg1_IsCaseInsensitive()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"stockholm\",1)";
                sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"STOCKHOLM\",1)";
                sheet.Cells["B3"].Formula = "REGEXTEST(A1,\"Stockholm\",1)";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["B1"].Value); // case ignored -> matches
                Assert.AreEqual(true, sheet.Cells["B2"].Value);
                Assert.AreEqual(true, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void CaseSensitivity_SwedishCharacters()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "ÖREBRO";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"örebro\",0)";
                sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"örebro\",1)";
                sheet.Calculate();

                Assert.AreEqual(false, sheet.Cells["B1"].Value); // case-sensitive, no match
                Assert.AreEqual(true, sheet.Cells["B2"].Value);  // case-insensitive, matches
            }
        }

        [TestMethod]
        public void CaseSensitivity_Omitted_DefaultsToCaseSensitive()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"stockholm\")";
                sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"STOCKHOLM\")";
                sheet.Cells["B3"].Formula = "REGEXTEST(A1,\"Stockholm\")";
                sheet.Calculate();

                Assert.AreEqual(false, sheet.Cells["B1"].Value); // default == case-sensitive
                Assert.AreEqual(false, sheet.Cells["B2"].Value);
                Assert.AreEqual(true, sheet.Cells["B3"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Invalid arguments
        // -------------------------------------------------------------------

        [TestMethod]
        public void InvalidCaseArgument_ReturnsValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"k\",2)";  // 2 is out of range
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void InvalidPattern_ReturnsValueError()
        {
            // Excel returns #VALUE! for syntactically invalid regex patterns.
            // The function must catch the regex exception and return a Value error.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"[\")";     // unterminated class
                sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"(\")";     // unterminated group
                sheet.Cells["B3"].Formula = "REGEXTEST(A1,\"*abc\")";  // quantifier without expression
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B3"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Numeric input is coerced to text. The coercion is culture sensitive,
        // so the test pins the culture (verified against Swedish Excel).
        // -------------------------------------------------------------------

        [TestMethod]
        public void NumericInput_IntegerMatchesPattern()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = 2026;
                sheet.Cells["A2"].Value = 1000000;

                sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"202[456]\")";
                sheet.Cells["B2"].Formula = "REGEXTEST(A2,\"^\\d+$\")";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["B1"].Value);
                Assert.AreEqual(true, sheet.Cells["B2"].Value); // no thousand separator in coercion
            }
        }

        [TestMethod]
        public void NumericInput_DecimalSeparatorIsCultureSensitive()
        {
            // Swedish culture renders 3.14 as "3,14", so a comma pattern matches
            // and a literal-dot pattern does not.
            SwitchToCulture("sv-SE");
            try
            {
                using (var package = OpenPackage("Testpackage"))
                {
                    var sheet = package.Workbook.Worksheets.Add("testsheet");
                    sheet.Cells["A1"].Value = 3.14;

                    sheet.Cells["B1"].Formula = "REGEXTEST(A1,\"3[.,]14\")"; // comma or dot
                    sheet.Cells["B2"].Formula = "REGEXTEST(A1,\"3\\.14\")";  // literal dot only
                    sheet.Calculate();

                    Assert.AreEqual(true, sheet.Cells["B1"].Value);
                    Assert.AreEqual(false, sheet.Cells["B2"].Value);
                }
            }
            finally
            {
                SwitchBackToCurrentCulture();
            }
        }

        // -------------------------------------------------------------------
        // Range / broadcast behavior
        // -------------------------------------------------------------------

        [TestMethod]
        public void RangeInput_PairwiseEqualDimensions()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "Linköping";
                sheet.Cells["A3"].Value = "Örebro";

                sheet.Cells["B1"].Value = "^S";       // matches Stockholm
                sheet.Cells["B2"].Value = "^Q";       // no match
                sheet.Cells["B3"].Value = "[A-ZÅÄÖ]"; // matches (starts with capital)

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A3,B1:B3)";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(false, sheet.Cells["D2"].Value);
                Assert.AreEqual(true, sheet.Cells["D3"].Value);
            }
        }

        [TestMethod]
        public void RangeInput_BroadcastAnchorIsTrue()
        {
            // 3 texts vs 2 patterns. Only the anchor cell is verified against Excel (TRUE).
            // The full spill (row 2 TRUE, row 3 #N/A) is expected but should be confirmed
            // once the round-2 verification workbook is filled in.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm 2026";
                sheet.Cells["A2"].Value = "Linköping 2025";
                sheet.Cells["A3"].Value = "Örebro 2024";

                sheet.Cells["C1"].Value = "\\d{4}";
                sheet.Cells["C2"].Value = "[A-ZÅÄÖ]\\w+";

                sheet.Cells["E1"].Formula = "REGEXTEST(A1:A3,C1:C2)";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["E1"].Value); // "Stockholm 2026" vs \d{4}
            }
        }

        [TestMethod]
        public void RangeInput_InvalidPattern_FailsPerCell()
        {
            // Excel isolates an invalid pattern to its own cell (#VALUE!) and still
            // computes the others. The current range loop lets the exception bubble out,
            // turning the WHOLE array into #VALUE!. The last row verifies the loop does
            // not stop at the error cell. Expected to FAIL until the loop catches per cell.
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

                sheet.Cells["D1"].Formula = "REGEXTEST(A1:A4,B1:B4)";
                sheet.Calculate();

                Assert.AreEqual(true, sheet.Cells["D1"].Value);
                Assert.AreEqual(true, sheet.Cells["D2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D3"].Value);
                Assert.AreEqual(true, sheet.Cells["D4"].Value); // computed AFTER the error cell
            }
        }
    }
}