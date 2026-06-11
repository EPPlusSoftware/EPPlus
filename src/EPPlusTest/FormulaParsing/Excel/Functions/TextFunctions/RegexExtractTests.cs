using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    /// <summary>
    /// Tests for the REGEXEXTRACT function. All expected values are verified against
    /// the calculation of Excel desktop (see REGEX verification workbook).
    /// </summary>
    [TestClass]
    public class RegexExtractTests : TestBase
    {
        // -------------------------------------------------------------------
        // return_mode 0 (first match, default)
        // -------------------------------------------------------------------

        [TestMethod]
        public void ReturnMode0_ReturnsFirstMatch()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1 b2 c3";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\w\\d\",0)";
                sheet.Calculate();

                Assert.AreEqual("a1", sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void ReturnMode0_NoMatch_ReturnsNA()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\d+\",0)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["B1"].Value);
            }
        }

        // -------------------------------------------------------------------
        // return_mode 1 (all matches, spills horizontally)
        // -------------------------------------------------------------------

        [TestMethod]
        public void ReturnMode1_ReturnsAllMatches()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Just #fitness finished 5k! #running";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"#\\w+\",1)";
                sheet.Calculate();

                Assert.AreEqual("#fitness", sheet.Cells["B1"].Value);
                Assert.AreEqual("#running", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ReturnMode1_NoMatch_ReturnsNA()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\d+\",1)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["B1"].Value);
            }
        }

        // -------------------------------------------------------------------
        // return_mode 2 (capturing groups, spills horizontally)
        // -------------------------------------------------------------------

        [TestMethod]
        public void ReturnMode2_ReturnsCapturingGroups()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "9183-Green-M";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"(\\d{4})-(\\w+)-(\\w+)\",2)";
                sheet.Calculate();

                Assert.AreEqual("9183", sheet.Cells["B1"].Value); // anchor verified against Excel
                Assert.AreEqual("Green", sheet.Cells["C1"].Value);
                Assert.AreEqual("M", sheet.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void ReturnMode2_NoMatch_ReturnsNA()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";
                sheet.Cells["A2"].Value = "abc";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"(\\d+)\",2)";
                sheet.Cells["B2"].Formula = "REGEXEXTRACT(A2,\"(\\d+)\",2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["B1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["B2"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Invalid return_mode
        // -------------------------------------------------------------------

        [TestMethod]
        public void ReturnMode3_ReturnsValueError()
        {
            // Excel: return_mode 3 is out of range and returns #VALUE!.
            // KNOWN BUG: the scalar branch validates "returnMode > 3" instead of "> 2",
            // so mode 3 currently falls through and returns the first match ("a1").
            // This test is expected to FAIL until the scalar validation is changed to "> 2".
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1 b2";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\w\\d\",3)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void ReturnModeNegative_ReturnsValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1 b2";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\w\\d\",-1)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }

        // -------------------------------------------------------------------
        // Invalid pattern
        // -------------------------------------------------------------------

        [TestMethod]
        public void InvalidPattern_ReturnsValueError()
        {
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "Stockholm";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"[\")";
                sheet.Cells["B2"].Formula = "REGEXEXTRACT(A1,\"(\")";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B2"].Value);
            }
        }

        [TestMethod]
        public void Range_Mode1_NoMatchCell_ReturnsNA()
        {
            // Range mode 1 yields the FIRST match per row; a non-matching row gives #N/A.
            // Current code calls .First() on an empty collection -> throws (whole array fails).
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "#fitness #running";
                sheet.Cells["A2"].Value = "#nature";
                sheet.Cells["A3"].Value = "Katt utan tagg";
                sheet.Cells["A4"].Value = "#a #b";
                sheet.Cells["B1"].Value = "#\\w+";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A4,B1,1)";
                sheet.Calculate();

                Assert.AreEqual("#fitness", sheet.Cells["D1"].Value);
                Assert.AreEqual("#nature", sheet.Cells["D2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D3"].Value);
                Assert.AreEqual("#a", sheet.Cells["D4"].Value); // first match only, per row
            }
        }

        [TestMethod]
        public void Range_Mode2_PatternWithoutGroups_ReturnsValueError()
        {
            // Mode 2 with a pattern that has no capturing groups -> #VALUE! per cell.
            // Current code calls Skip(1).First() on an empty collection -> throws.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "9183";
                sheet.Cells["A2"].Value = "abcd";
                sheet.Cells["B1"].Value = "\\d+"; // no capturing group

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2,B1,2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void Range_Mode2_GroupsButNoMatchCell_ReturnsNA()
        {
            // Mode 2 with groups: matching row gives the first group, non-matching row gives #N/A.
            // Current code returns an empty string for the non-matching row (no throw).
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "9183-Green-M";
                sheet.Cells["A2"].Value = "ingen match";
                sheet.Cells["B1"].Value = "(\\d{4})-(\\w+)-(\\w+)";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2,B1,2)";
                sheet.Calculate();

                Assert.AreEqual("9183", sheet.Cells["D1"].Value); // first group only, in range mode
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void Range_NegativeReturnMode_ReturnsValueError()
        {
            // Range branch validates with Math.Abs, so -1/-2 slip through into mode 0.
            // Excel returns #VALUE! per cell, as the scalar branch already does.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1 b2";
                sheet.Cells["A2"].Value = "c3 d4";
                sheet.Cells["B1"].Value = "\\w\\d";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2,B1,-1)";
                sheet.Cells["F1"].Formula = "REGEXEXTRACT(A1:A2,B1,-2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D2"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["F1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["F2"].Value);
            }
        }

        [TestMethod]
        public void Range_NegativeCaseSensitivity_ReturnsValueError()
        {
            // caseSensitivity -1 passes the Math.Abs check and reaches (RegexOptions)(-1).
            // Excel returns #VALUE! per cell.
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "a1 b2";
                sheet.Cells["A2"].Value = "c3 d4";
                sheet.Cells["B1"].Value = "\\w\\d";

                sheet.Cells["D1"].Formula = "REGEXEXTRACT(A1:A2,B1,0,-1)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D1"].Value);
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void ReturnMode2_PatternWithoutGroups_ReturnsValueError()
        {
            // Mode 2 with a pattern that has no capturing groups returns #VALUE!,
            // both scalar and in range mode (verified against Excel).
            using (var package = OpenPackage("Testpackage"))
            {
                var sheet = package.Workbook.Worksheets.Add("testsheet");
                sheet.Cells["A1"].Value = "9183";

                sheet.Cells["B1"].Formula = "REGEXEXTRACT(A1,\"\\d+\",2)";
                sheet.Calculate();

                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["B1"].Value);
            }
        }
    }
}