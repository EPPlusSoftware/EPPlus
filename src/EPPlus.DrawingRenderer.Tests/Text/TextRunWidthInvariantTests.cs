using System.Globalization;
using System.Text;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;

namespace EPPlus.DrawingRenderer.Tests.Text
{
    /// <summary>
    /// Step 0 of the SVG font subsetting plan.
    ///
    /// SVG output will set textLength on every tspan, which assumes that run.Width is the width
    /// of exactly run.CurrentText in the run's font. These tests check that invariant for the runs
    /// produced by ParagraphRenderItem.WrapTextFragmentsAndGenerateTextRuns, and that glyph
    /// fallback (emoji) is part of the measured width.
    /// </summary>
    [TestClass]
    public class TextRunWidthInvariantTests : FontTestBase
    {
        private const string Roboto = "Roboto";
        private const string OpenSans = "Open Sans";

        // Known issue, handed over: when a line ends, trailing spaces are trimmed from the line's
        // text but their width stays in the last run. Not visible today (nothing is placed after the
        // last run on a line), but textLength in SVG step 5 needs it fixed. Remove the Ignore
        // attributes when it is.
        private const string TrailingSpaceIssue =
            "Known issue: trailing spaces at line ends are included in the last run's width but not in its text. " +
            "Handed over (see 'Fynd i textradbrytningen'). Remove when fixed; required before textLength in SVG step 5.";

        // Widths are summed per character in doubles and compared with a single shaping pass in
        // floats; anything above this is a real difference, not rounding.
        private const double Tolerance = 0.01;

        // Kerning is attributed to the left glyph of a pair. When a line breaks inside a word, the
        // kern value of the pair at the break stays with the run on the first line. Allow for one
        // kern value at 11pt.
        private const double MidWordBreakTolerance = 0.6;

        private const float FontSize = 11f;

        // Test fonts only, so the measurements are the same on every machine.
        private static RenderContext Rc
        {
            get { return TestFolderContext; }
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            // The test fonts are part of the test setup, so a missing font is a failure, not
            // an inconclusive result.
            AssertTestFont(Roboto, FontSubFamily.Regular);
            AssertTestFont(Roboto, FontSubFamily.Bold);
            AssertTestFont(OpenSans, FontSubFamily.Regular);
        }

        // -----------------------------------------------------------------------------------------
        // Width invariant
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_WrapInsideFragment()
        {
            var fragments = new[] { Rt("The quick brown fox jumps over the lazy dog") };
            var p = CreateParagraph(80, fragments);

            Assert.IsTrue(p.Lines.Count > 1, "Expected the text to wrap.");
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_WrapExactlyAtFragmentBoundary()
        {
            var first = Rt("Hello ");
            var second = Rt("World again");

            // Wide enough for "Hello " but not for "Hello W", so the break falls on the boundary.
            var widthHelloSpace = Measure("Hello ", first);
            var widthHelloW = Measure("Hello W", first);
            var maxWidth = Math.Ceiling(widthHelloSpace) + 1;
            Assert.IsTrue(maxWidth < widthHelloW, "Test setup: max width must not fit 'Hello W'.");

            var fragments = new[] { first, second };
            var p = CreateParagraph(maxWidth, fragments);

            Assert.IsTrue(p.Lines.Count > 1, "Expected the text to wrap.");
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_TrailingSpace_NoWrap()
        {
            var fragments = new[] { Rt("Hello world ") };
            var p = CreateParagraph(1000, fragments);

            Assert.AreEqual(1, p.Lines.Count);
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        public void Invariant_DoubleSpaces_NoWrap()
        {
            var fragments = new[] { Rt("Hello  world  again") };
            var p = CreateParagraph(1000, fragments);

            Assert.AreEqual(1, p.Lines.Count);
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_DoubleSpaces_Wrapped()
        {
            var fragments = new[] { Rt("alpha  beta  gamma  delta  epsilon") };
            var p = CreateParagraph(60, fragments);

            Assert.IsTrue(p.Lines.Count > 1, "Expected the text to wrap.");
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_ExplicitLineBreak()
        {
            var fragments = new[] { Rt("First line \r\nSecond line") };
            var p = CreateParagraph(1000, fragments);

            Assert.AreEqual(2, p.Lines.Count);
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_TwoFontsInParagraph()
        {
            var fragments = new[]
            {
                Rt("Alpha beta gamma "),
                Rt("delta epsilon zeta ", Roboto, bold: true),
                Rt("eta theta iota kappa", OpenSans),
            };
            var p = CreateParagraph(90, fragments);

            Assert.IsTrue(p.Lines.Count > 1, "Expected the text to wrap.");
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        [Ignore(TrailingSpaceIssue)]
        public void Invariant_HyphenAsAdditionalBreakChar()
        {
            const string text = "well-known long-term state-of-the-art";
            var fragments = new[] { Rt(text) };
            var p = CreateParagraph(50, fragments);

            p.AddAdditionalWrappingChars(new[] { '-' });
            p.SetNewText(text); // re-wraps with the additional break char

            Assert.IsTrue(p.Lines.Count > 1, "Expected the text to wrap.");
            AssertRunWidthsMatchShaping(p, fragments);
        }

        [TestMethod]
        public void Invariant_BreakInsideWordWithoutSpace()
        {
            var fragments = new[] { Rt("Supercalifragilisticexpialidocious") };
            var p = CreateParagraph(40, fragments);

            Assert.IsTrue(p.Lines.Count > 1, "Expected the word to be broken.");
            AssertRunWidthsMatchShaping(p, fragments, MidWordBreakTolerance);
        }

        // -----------------------------------------------------------------------------------------
        // Glyph fallback in measurement
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void GlyphFallback_EmojiIsMeasuredWithNotoEmojiAdvance()
        {
            const uint chartIncreasing = 0x1F4C8; // 📈
            var emoji = char.ConvertFromUtf32((int)chartIncreasing);

            var roboto = TestFolderEngine.LoadFont(Roboto);
            Assert.IsFalse(roboto.CmapTable.TryGetGlyphId(chartIncreasing, out _), "Test setup: Roboto must not have the emoji.");

            // Where the provider routes the emoji, and that glyph's advance in points.
            var provider = new DefaultFontProvider(TestFolderEngine, roboto);
            Assert.IsTrue(provider.TryGetGlyphFont(chartIncreasing, out var fallbackFont, out var glyphId), "No fallback font has the emoji.");
            Assert.AreNotSame(roboto, fallbackFont);
            StringAssert.Contains(fallbackFont.GetEnglishFontFamilyName(), "Emoji");
            var emojiAdvance = fallbackFont.HmtxTable.GetAdvanceWidth(glyphId) * (double)FontSize / fallbackFont.HeadTable.UnitsPerEm;
            var notdefAdvance = roboto.HmtxTable.GetAdvanceWidth(0) * (double)FontSize / roboto.HeadTable.UnitsPerEm;
            Assert.AreNotEqual(notdefAdvance, emojiAdvance, 0.001, "Test setup: emoji and .notdef advances must differ.");

            // The width the layout path measured for the run.
            var fragments = new[] { Rt("Sales " + emoji) };
            var p = CreateParagraph(1000, fragments);
            Assert.AreEqual(1, p.Runs.Count);

            var expected = Measure("Sales ", fragments[0]) + emojiAdvance;
            Assert.AreEqual(expected, p.Runs[0].Width, Tolerance,
                "The emoji should be measured with the fallback font's advance, not Roboto's .notdef.");

            AssertRunWidthsMatchShaping(p, fragments);
        }

        // -----------------------------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------------------------

        private static void AssertTestFont(string family, FontSubFamily subFamily)
        {
            Assert.AreEqual(FontAvailability.Exact, TestFolderEngine.GetFontAvailability(family, subFamily),
                $"Test font '{family} {subFamily}' is not available in {FontFolder}.");
        }

        private static RichTextFormatSimple Rt(string text, string family = Roboto, bool bold = false)
        {
            return new RichTextFormatSimple(text, family, FontSize, bold);
        }

        /// <summary>
        /// Creates an auto-sized paragraph that wraps at <paramref name="maxWidth"/> points.
        /// Auto-size is used because a fixed-size paragraph shrinks its own width to the widest
        /// line after the first wrap, and later fragments would then wrap at that width.
        /// </summary>
        private static ParagraphRenderItem CreateParagraph(double maxWidth, params RichTextFormatSimple[] fragments)
        {
            var group = new GroupRenderItem(new BoundingBox(0, 0, 1000, 1000));
            var body = new SvgTextBodyRenderItem(Rc, group, 0, 0, maxWidth, 1000, false, autoSize: true);

            var p = body.AddParagraph(fragments[0]);
            for (int i = 1; i < fragments.Length; i++)
            {
                p.AddRichText(fragments[i]);
            }
            return p;
        }

        private static double Measure(string text, IFontFormatBase font)
        {
            return Rc.GetMeasurementShaper(font).ShapeLight(text).GetWidthInPoints(font.Size);
        }

        /// <summary>
        /// Asserts that every run's width equals the width of its own text shaped in its own font.
        /// The font is taken from the input fragment the run came from (OriginalRtIdx), since
        /// ParagraphRenderItem.AddRichText(IRichTextFormatSimple) does not re-import styles to the
        /// runs it recreates. All mismatches are reported together.
        /// </summary>
        private static void AssertRunWidthsMatchShaping(ParagraphRenderItem p, RichTextFormatSimple[] fragments, double tolerance = Tolerance)
        {
            Assert.IsTrue(p.Runs.Count > 0, "The paragraph has no runs.");

            var report = new StringBuilder();
            var failures = 0;

            foreach (var run in p.Runs)
            {
                if (string.IsNullOrEmpty(run.CurrentText))
                    continue;

                Assert.IsTrue(run.OriginalRtIdx >= 0 && run.OriginalRtIdx < fragments.Length,
                    $"Run '{run.CurrentText}' has no valid source fragment (OriginalRtIdx {run.OriginalRtIdx}).");
                var font = fragments[run.OriginalRtIdx];

                var expected = Measure(run.CurrentText, font);
                var diff = run.Width - expected;
                var ok = Math.Abs(diff) <= tolerance;
                if (!ok) failures++;

                report.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "{0} '{1}' [{2} {3}]: run.Width={4:F3} shaped={5:F3} diff={6:+0.000;-0.000;0.000}",
                    ok ? "  " : "!!", run.CurrentText, font.Family, font.SubFamily, run.Width, expected, diff));
            }

            if (failures > 0)
            {
                Assert.Fail($"{failures} run(s) differ from their shaped width by more than {tolerance}pt:{Environment.NewLine}{report}");
            }
        }
    }
}
