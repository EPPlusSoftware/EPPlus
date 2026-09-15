using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
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

        private static TextLayoutEngine CreateEngine()
        {
            Assert.Inconclusive("Fill in TextLayoutEngine construction before running these tests.");
            return null;
        }

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
    }
}
