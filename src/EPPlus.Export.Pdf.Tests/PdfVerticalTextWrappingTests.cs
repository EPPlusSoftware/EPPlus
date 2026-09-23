using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class PdfVerticalTextWrappingTests : PdfTestBase
    {
        private const string FontName = "Aptos Narrow";
        private const float FontSize = 11;
        private OpenTypeFontEngine _fontEngine;

        [TestInitialize]
        public void Init()
        {
            _fontEngine = new OpenTypeFontEngine();
            _fontEngine.RequireExactFont = true;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _fontEngine?.Dispose();
        }

        /// <summary>
        /// WrapRichTextLineCollectionVertical runs FinalizeLineFragments twice on each line:
        /// once in WrapRichTextLines' ascent/descent pass, and once more in the
        /// TextLineCollection constructor (finalizeLineFragments: true). Without a Clear()
        /// at the top of FinalizeLineFragments the output fragments accumulate, which also
        /// defeats the count-equality guard in FinalizeTextLineData.
        /// </summary>
        [TestMethod]
        public void VerticalWrap_LineFragmentsDoNotAccumulateOnDoubleFinalize()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);

            var collection = engine.WrapRichTextLineCollectionVertical(
                FragmentsTyped("aaa bbb ccc"), step * 8);

            Assert.AreEqual(2, collection.Count);
            foreach (var line in collection)
            {
                Assert.AreEqual(line.InternalLineFragments.Count, line.LineFragments.Count,
                    "LineFragments must not accumulate across repeated finalization.");
            }
        }
        /// <summary>
        /// THE discriminating test. Narrow and wide glyphs must wrap at the exact same
        /// character count, because vertical stacking ignores glyph advance entirely.
        /// If this fails, the code is still measuring horizontal widths somewhere.
        /// </summary>
        [TestMethod]
        public void VerticalWrap_IsIndependentOfGlyphWidth()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);
            var max = step * 5;

            var narrow = engine.WrapVerticalRichTextLines(Fragments("iiiiiiiiii"), max);
            var wide = engine.WrapVerticalRichTextLines(Fragments("mmmmmmmmmm"), max);

            Assert.AreEqual(narrow.Count, wide.Count, "Stack count must not depend on glyph width.");
            Assert.AreEqual(2, narrow.Count, "10 chars at 5 per stack should yield 2 stacks.");
            CollectionAssert.AreEqual(
                narrow.Select(l => l.Text.Length).ToList(),
                wide.Select(l => l.Text.Length).ToList(),
                "Break positions must be identical for narrow and wide glyphs.");
        }

        [TestMethod]
        public void VerticalWrap_BreaksOnLastSpaceThatFits()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);

            var lines = engine.WrapVerticalRichTextLines(Fragments("aaa bbb ccc"), step * 8);

            Assert.AreEqual(2, lines.Count);
            Assert.AreEqual("aaa bbb", lines[0].Text, "Trailing space must be trimmed from the stack text.");
            Assert.AreEqual("ccc", lines[1].Text);
            Assert.IsTrue(lines[0].WasWrappedOnSpace, "Break on whitespace should be flagged.");
            Assert.IsLessThan(lines[0].LargestAscent, 0);
            Assert.AreEqual(lines[0].InternalLineFragments.Count, lines[0].LineFragments.Count,
    "LineFragments must not accumulate across repeated finalization.");
        }

        [TestMethod]
        public void HorizontalWrap_StillDependsOnGlyphWidth()
        {
            var engine = CreateEngine();

            var narrow = engine.WrapRichTextLines(Fragments("iiiiiiiiiiiiiiiiiiii"), 50d, false);
            var wide = engine.WrapRichTextLines(Fragments("mmmmmmmmmmmmmmmmmmmm"), 50d, false);

            Assert.IsTrue(wide.Count > narrow.Count,
                "Horizontal wrapping must still measure glyph advances.");
        }

        private TextLayoutEngine CreateEngine()
            => _fontEngine.GetTextLayoutEngine(FontName, FontSubFamily.Regular);

        private static List<ITextFragmentBase> Fragments(params string[] texts)
        {
            var list = new List<ITextFragmentBase>(texts.Length);
            foreach (var text in texts)
            {
                var frag = new TextFragment();
                frag.Font = new RichTextFormatSimple();
                frag.Text = text;
                frag.Font.Family = FontName;
                frag.Font.Size = FontSize;
                list.Add(frag);
            }
            return list;
        }

        /// <summary>
        /// Runs an unconstrained pass to learn the per-character step for this font/size.
        /// AscentPoints/DescentPoints are populated on the fragment during ProcessFragment,
        /// so the tests never have to hardcode a point value that depends on font metrics.
        /// </summary>
        private static double GetStep(TextLayoutEngine engine)
        {
            var frags = Fragments("X");
            engine.WrapVerticalRichTextLines(frags, double.MaxValue);
            return frags[0].AscentPoints + frags[0].DescentPoints;
        }
        private static List<TextFragment> FragmentsTyped(params string[] texts)
        {
            var list = new List<TextFragment>(texts.Length);
            foreach (var text in texts)
            {
                var frag = new TextFragment();
                frag.Font = new RichTextFormatSimple();
                frag.Text = text;
                frag.Font.Family = FontName;
                frag.Font.Size = FontSize;
                list.Add(frag);
            }
            return list;
        }
    }
}
