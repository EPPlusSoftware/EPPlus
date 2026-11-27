using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class NumberFitFormatterTests
    {
        private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");
        private static readonly CultureInfo SvSe = CultureInfo.GetCultureInfo("sv-SE");

        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        private CultureInfo _prevCulture;

        [TestInitialize]
        public void Init()
        {
            _prevCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = EnUs;
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            Thread.CurrentThread.CurrentCulture = _prevCulture;
            _sheet = null;
            _package.Dispose();
            _package = null;
        }

        #region Helpers
        //A new translator per call, as GetFormatPart mutates the General format part.
        private static ExcelFormatTranslator Nf(string format) => new ExcelFormatTranslator(format, -1);

        private static List<string> Candidates(object value, string format, CultureInfo ci = null)
            => NumberFitFormatter.GetCandidates(value, Nf(format), ci ?? EnUs).ToList();

        /// <summary>
        /// Mirrors the selection loop in DrawingTextBody, with the character count as the "width".
        /// </summary>
        private static string Fit(object value, string format, int maxChars)
        {
            string chosen = null;
            foreach (var s in NumberFitFormatter.GetCandidates(value, Nf(format), EnUs))
            {
                chosen = s;
                if (s.Length <= maxChars) break;
            }
            return chosen;
        }
        #endregion

        #region Same output as ExcelRangeBase.Text
        [DataTestMethod]
        [DataRow(1234.5678, "General")]
        [DataRow(-1234.5678, "General")]
        [DataRow(0.000123456789, "General")]
        [DataRow(123456789012d, "General")]
        [DataRow(1234.5678, "0.000")]
        [DataRow(1234.5678, "#,##0.00")]
        [DataRow(0.1234, "0.0%")]
        [DataRow(-1234.5678, "#,##0.00;(#,##0.00)")]
        [DataRow(45000.75, "yyyy-mm-dd")]
        public void FirstCandidateEqualsRangeText(double value, string format)
        {
            var cell = _sheet.Cells["A1"];
            cell.Value = value;
            cell.Style.Numberformat.Format = format;

            var nf = ValueToTextHandler.GetNumberFormat(cell.StyleID, _package.Workbook.Styles).FormatTranslator;
            var first = NumberFitFormatter.GetCandidates(value, nf, null).First();

            Assert.AreEqual(cell.Text, first);
        }
        #endregion

        #region General - fixed notation
        [TestMethod]
        public void GeneralReducesDecimalsOneAtATime()
        {
            var c = Candidates(1234.5678, "General");

            CollectionAssert.AreEqual(
                new[] { "1234.5678", "1234.568", "1234.57", "1234.6", "1235" },
                c.Take(5).ToArray());
        }

        [TestMethod]
        public void GeneralNegativeValue()
        {
            var c = Candidates(-1234.5678, "General");

            CollectionAssert.AreEqual(
                new[] { "-1234.5678", "-1234.568", "-1234.57", "-1234.6", "-1235" },
                c.Take(5).ToArray());
        }

        [DataTestMethod]
        [DataRow(2.675, "2.68")]
        [DataRow(-2.675, "-2.68")]
        [DataRow(39.285, "39.29")]
        [DataRow(1.005, "1.01")]
        public void GeneralRoundsMidpointAwayFromZeroLikeExcel(double value, string expected)
        {
            var c = Candidates(value, "General");
            CollectionAssert.Contains(c, expected);
        }

        [TestMethod]
        public void GeneralValueBelowOneEndsWithZero()
        {
            var c = Candidates(0.000123456789, "General");

            CollectionAssert.Contains(c, "0.00012");
            Assert.AreEqual("0", c.Last());
            Assert.IsFalse(c.Any(x => x.Contains("E")), "No scientific candidates after reaching 0.");
        }

        [TestMethod]
        public void GeneralWithSuffixKeepsSuffix()
        {
            var c = Candidates(1234.5678, "General\" kr\"");

            CollectionAssert.Contains(c, "1234.57 kr");
            CollectionAssert.Contains(c, "1235 kr");
        }

        [TestMethod]
        public void GeneralUsesCultureDecimalSeparator()
        {
            var c = Candidates(1234.5678, "General", SvSe);

            CollectionAssert.Contains(c, "1234,57");
            Assert.IsFalse(c.Any(x => x.Contains(".")), "en-US separator should not be used.");
        }
        #endregion

        #region General - scientific notation
        [TestMethod]
        public void GeneralFallsBackToScientificAfterInteger()
        {
            var c = Candidates(123456789d, "General");

            var ix = c.IndexOf("123456789");
            Assert.IsTrue(ix >= 0);
            CollectionAssert.AreEqual(
                new[] { "1.23457E+08", "1.2346E+08", "1.235E+08", "1.23E+08", "1.2E+08", "1E+08" },
                c.Skip(ix + 1).ToArray());
        }

        [TestMethod]
        public void GeneralThatIsAlreadyScientificOnlyGivesScientificCandidates()
        {
            var c = Candidates(1.23456789012E+15, "General");

            Assert.IsTrue(c[0].Contains("E"), $"Expected .Text to be scientific, was {c[0]}");
            Assert.IsTrue(c.All(x => x.Contains("E")));
            Assert.AreEqual("1E+15", c.Last());
        }
        #endregion

        #region Ordering and duplicates
        [DataTestMethod]
        [DataRow(1234.5678)]
        [DataRow(-98765.4321)]
        [DataRow(0.000123456789)]
        [DataRow(123456789d)]
        [DataRow(3.14159265358979)]
        public void CandidatesAreDistinct(double value)
        {
            var c = Candidates(value, "General");
            Assert.AreEqual(c.Count, c.Distinct().Count());
        }

        [DataTestMethod]
        [DataRow(1234.5678)]
        [DataRow(-98765.4321)]
        [DataRow(3.14159265358979)]
        public void FixedCandidatesNeverGetLonger(double value)
        {
            var fixedPart = Candidates(value, "General").TakeWhile(x => x.Contains("E") == false).ToList();
            for (int i = 1; i < fixedPart.Count; i++)
            {
                Assert.IsTrue(fixedPart[i].Length <= fixedPart[i - 1].Length,
                    $"'{fixedPart[i]}' is longer than '{fixedPart[i - 1]}'");
            }
        }
        #endregion

        #region Values and formats that are not reduced
        [DataTestMethod]
        [DataRow(1234.5678, "0.000", "1234.568")]
        [DataRow(1234.5678, "#,##0.00", "1,234.57")]
        [DataRow(0.1234, "0.0%", "12.3%")]
        public void ExplicitFormatGivesSingleCandidate(double value, string format, string expected)
        {
            var c = Candidates(value, format);
            CollectionAssert.AreEqual(new[] { expected }, c.ToArray());
        }

        [TestMethod]
        public void DateFormatGivesSingleCandidate()
        {
            var c = Candidates(45000d, "yyyy-mm-dd");
            Assert.AreEqual(1, c.Count);
        }

        [TestMethod]
        public void StringValueGivesSingleCandidate()
        {
            var c = Candidates("Text", "General");
            CollectionAssert.AreEqual(new[] { "Text" }, c.ToArray());
        }

        [TestMethod]
        public void BoolValueGivesSingleCandidate()
        {
            var c = Candidates(true, "General");
            CollectionAssert.AreEqual(new[] { "TRUE" }, c.ToArray());
        }

        [TestMethod]
        public void ZeroGivesSingleCandidate()
        {
            var c = Candidates(0d, "General");
            CollectionAssert.AreEqual(new[] { "0" }, c.ToArray());
        }

        [DataTestMethod]
        [DataRow(1234)]
        [DataRow(-42)]
        public void IntegerTypesAreHandled(int value)
        {
            var c = Candidates(value, "General");
            Assert.AreEqual(value.ToString(EnUs), c[0]);
        }
        #endregion

        #region Fitting (selection loop)
        [DataTestMethod]
        [DataRow(20, "1234.5678")]
        [DataRow(9, "1234.5678")]
        [DataRow(8, "1234.568")]
        [DataRow(6, "1234.6")]
        [DataRow(4, "1235")]
        [DataRow(3, "1E+03")]     //Nothing fits. Use the last (shortest) candidate.
        public void FitPicksFirstCandidateThatFits(int maxChars, string expected)
        {
            Assert.AreEqual(expected, Fit(1234.5678, "General", maxChars));
        }

        [TestMethod]
        public void FitNeverTruncatesExplicitFormat()
        {
            //The old data table code removed characters from the end. 1234.568 must not become 123.
            Assert.AreEqual("1234.568", Fit(1234.5678, "0.000", 3));
        }
        #endregion

        #region RoundingHelper.RoundToDecimals
        [DataTestMethod]
        [DataRow(2.675, 2, 2.68)]
        [DataRow(-2.675, 2, -2.68)]
        [DataRow(39.285, 2, 39.29)]
        [DataRow(1.005, 2, 1.01)]
        [DataRow(0.5, 0, 1d)]
        [DataRow(-0.5, 0, -1d)]
        [DataRow(1234.5678, 0, 1235d)]
        [DataRow(1234.5678, 3, 1234.568)]
        [DataRow(0.000123456789, 5, 0.00012)]
        public void RoundToDecimals(double value, int decimals, double expected)
        {
            Assert.AreEqual(expected, RoundingHelper.RoundToDecimals(value, decimals));
        }

        [TestMethod]
        public void RoundToDecimalsOutsideDecimalRangeReturnsInput()
        {
            Assert.AreEqual(1E+30, RoundingHelper.RoundToDecimals(1E+30, 2));
            Assert.AreEqual(-1E+30, RoundingHelper.RoundToDecimals(-1E+30, 2));
        }
        #endregion
    }
}
