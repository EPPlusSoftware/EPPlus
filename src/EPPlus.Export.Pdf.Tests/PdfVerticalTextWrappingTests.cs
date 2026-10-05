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
using System.Diagnostics;
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

        [TestMethod]
        public void VerticalTextWrappingBasicTest()
        {
            using (var package = OpenTemplatePackage("wrappingVerticalTExtPdf.xlsx"))
            {
                var ws = package.Workbook.Worksheets[1];
                var path = _pdfPath + "wrappingVerticalTExtPdf.pdf";
                ws.SaveAsPdf(path);
            }
        }

        [TestMethod]
        public void VerticalTextTestSheet1() 
        {
            using (var package = OpenTemplatePackage("TestsVerticalText.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                var path = _pdfPath + "verticalTextRegression.pdf";
                ws.SaveAsPdf(path);
            }
        }


        [TestMethod]
        public void VerticalTextTestSheetWrapping()
        {
            using (var package = OpenTemplatePackage("TestsVerticalText.xlsx"))
            {
                var ws = package.Workbook.Worksheets[6];
                var path = _pdfPath + "mergeRegression.pdf";
                ws.SaveAsPdf(path);
            }
        }

        [TestMethod]
        public void VerticalTextTestSheet2()
        {
            using (var package = OpenTemplatePackage("TestsVerticalText.xlsx"))
            {
                var ws = package.Workbook.Worksheets[1];
                var path = _pdfPath + "verticalTextRegressionSheet2.pdf";
                ws.SaveAsPdf(path);
            }
        }

        [TestMethod]
        public void VerticalTextTestSheet3()
        {
            using (var package = OpenTemplatePackage("TestsVerticalText.xlsx"))
            {
                var ws = package.Workbook.Worksheets[3];
                var path = _pdfPath + "wrappedHorizontal.pdf";
                ws.SaveAsPdf(path);
            }
        }

        [TestMethod] 
        public void DumpAptosNarrowMetrics()
        {
            using (var engine = new OpenTypeFontEngine())
            {
                var font = engine.LoadFont("Aptos", FontSubFamily.Regular);
                double upem = font.HeadTable.UnitsPerEm;

                void Show(string name, double units)
                {
                    Debug.WriteLine($"{name,-40} {units,8:F0} units = {units / upem,7:F5} em " +
                                    $"= {units / upem * 11.04,7:F3} pt @11.04 " +
                                    $"= {units / upem * 11,7:F3} pt @11");
                }

                Show("hhea asc - desc", font.HheaTable.ascender - font.HheaTable.descender);
                Show("hhea asc - desc + lineGap", font.HheaTable.ascender - font.HheaTable.descender + font.HheaTable.lineGap);
                Show("OS/2 typo asc - desc", font.Os2Table.sTypoAscender - font.Os2Table.sTypoDescender);
                Show("OS/2 typo asc - desc + gap", font.Os2Table.sTypoAscender - font.Os2Table.sTypoDescender + font.Os2Table.sTypoLineGap);
                Show("OS/2 win asc + win desc", font.Os2Table.usWinAscent + font.Os2Table.usWinDescent);
                Show("usWinAscent", font.Os2Table.usWinAscent);
                Show("usWinDescent", font.Os2Table.usWinDescent);
                Show("head yMax - yMin", font.HeadTable.Ymax - font.HeadTable.Ymin);
                Show("TARGET", 0);
                Debug.WriteLine($"Excel measured: 15.240 pt");
            }
        }

        /// <summary>
        /// 2.2 - verified against Excel: "1,2346E+19" in a 60.02pt row (capacity 3 chars)
        /// produced exactly "1,2" / "346" / "E+1" / "9".
        /// </summary>
        [TestMethod]
        public void VerticalWrap_NoSpaces_BreaksAtCapacity()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);

            var lines = engine.WrapVerticalRichTextLines(Fragments("1,2346E+19"), step * 3);

            Assert.AreEqual(4, lines.Count);
            Assert.AreEqual("1,2", lines[0].Text);
            Assert.AreEqual("346", lines[1].Text);
            Assert.AreEqual("E+1", lines[2].Text);
            Assert.AreEqual("9", lines[3].Text);
        }

        /// <summary>
        /// 2.5 - an explicit line break starts a new stack even when room remains.
        /// </summary>
        [TestMethod]
        public void VerticalWrap_ExplicitLineBreakStartsNewStack()
        {
            var engine = CreateEngine();

            var lines = engine.WrapVerticalRichTextLines(Fragments("ab\ncd"), double.MaxValue);

            Assert.AreEqual(2, lines.Count);
            Assert.AreEqual("ab", lines[0].Text);
            Assert.AreEqual("cd", lines[1].Text);
        }

        /// <summary>
        /// 2.6 - text that fits needs no break, and must not differ from the unwrapped path.
        /// </summary>
        [TestMethod]
        public void VerticalWrap_TextThatFitsProducesOneStack()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);

            var lines = engine.WrapVerticalRichTextLines(Fragments("Hej"), step * 10);

            Assert.AreEqual(1, lines.Count);
            Assert.AreEqual("Hej", lines[0].Text);
        }

        /// <summary>
        /// 2.4 - a row shorter than one line height. Every character immediately exceeds the
        /// limit, so WrapCurrentLine's else-branch fires on each one. The timeout guards the
        /// failure mode where the overflow character is never consumed.
        /// </summary>
        [TestMethod, Timeout(5000)]
        public void VerticalWrap_RowShorterThanOneStep_Terminates()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);

            var lines = engine.WrapVerticalRichTextLines(Fragments("abcde"), step * 0.5);

            Assert.AreEqual("abcde", string.Concat(lines.Select(l => l.Text)),
                "No characters may be lost when every character overflows.");
        }

        /// <summary>
        /// 2.7 - an empty fragment is skipped in the loop, leaving a TextLine with no
        /// InternalLineFragments. FinalizeLineFragments calls InternalLineFragments.Last()
        /// unguarded, so this is where it would throw.
        /// </summary>
        [TestMethod]
        public void VerticalWrap_EmptyText_DoesNotThrow()
        {
            var engine = CreateEngine();

            var lines = engine.WrapRichTextLines(Fragments(""), 100d, false);

            Assert.AreEqual(0, lines.Count);
        }

        /// <summary>
        /// 2.3 - a long run with no break opportunity. The point is that nothing is lost or
        /// duplicated when every stack breaks mid-word.
        /// </summary>
        [TestMethod, Timeout(5000)]
        public void VerticalWrap_LongWordWithoutSpaces_LosesNothing()
        {
            var engine = CreateEngine();
            var step = GetStep(engine);
            const string text = "abcdefghijklmnopqrstuvwxyzabcdefghijklmn";

            var lines = engine.WrapVerticalRichTextLines(Fragments(text), step * 7);

            Assert.AreEqual(text, string.Concat(lines.Select(l => l.Text)));
            Assert.IsTrue(lines.Count >= 6, "40 characters at 7 per stack needs at least 6 stacks.");
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
        /// Runs an unconstrained pass to learn the pe  r-character step for this font/size.
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
