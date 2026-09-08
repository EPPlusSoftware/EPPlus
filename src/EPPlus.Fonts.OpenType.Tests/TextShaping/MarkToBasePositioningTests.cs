/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Mark-to-base offset convention
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.TextShaping.Positioning;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Tests the offset convention for mark-to-base positioning (GPOS type 4).
    ///
    /// The offsets written by <c>MarkToBaseProvider</c> are relative to the base glyph's ORIGIN
    /// today. A renderer draws the mark at the pen, which has already moved by the base glyph's
    /// advance, so a base-relative offset is unusable without the renderer knowing the advance.
    /// The offsets are therefore being changed to PEN-RELATIVE, following the HarfBuzz convention.
    ///
    /// Every expected value below is taken straight from the anchors in the shipped test fonts:
    ///   pen-relative XOffset = baseAnchor.X - markAnchor.X - baseAdvance
    ///   YOffset              = baseAnchor.Y - markAnchor.Y
    ///
    /// The X assertions FAIL before the change and pass after it. The Y assertions pass both
    /// before and after - vertical position is unaffected by the pen convention, and they are
    /// here as guards against the change going too far.
    /// </summary>
    [TestClass]
    public class MarkToBasePositioningTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        /// <summary>
        /// Noto Sans Math has mark as lookup type 4 only, so nothing is silently skipped by the
        /// missing type 5 loader. Units per em is 1000.
        /// </summary>
        private const string NotoSansMathFamily = "Noto Sans Math";

        /// <summary>
        /// Mulish has both type 4 and type 5 mark lookups. Only type 4 is loaded, which makes it
        /// a useful second font: the expected values must come from the type 4 anchors alone.
        /// </summary>
        private const string MulishFamily = "Mulish";

        private const char CombiningAcute = '\u0301';

        // Noto Sans Math: A advance 639, base anchor (318, 714); a advance 561, base anchor
        // (281, 536); mark anchor (93, 536).
        private const int NotoMathBaseRelativeX_A = 225;
        private const int NotoMathPenRelativeX_A = -414;
        private const int NotoMathYOffset_A = 178;
        private const int NotoMathPenRelativeX_a = -373;
        private const int NotoMathYOffset_a = 0;

        // Mulish: A advance 735, base anchor (366, 705); a advance 600, base anchor (294, 484);
        // mark anchor (0, 484).
        private const int MulishPenRelativeX_A = -369;
        private const int MulishYOffset_A = 221;
        private const int MulishPenRelativeX_a = -306;
        private const int MulishYOffset_a = 0;

        #region Offsets are pen-relative

        [TestMethod]
        public void MarkOffset_OverCapital_IsPenRelative_NotoSansMath()
        {
            var glyphs = ShapeBaseAndMark(NotoSansMathFamily, 'A');

            Assert.AreEqual(
                NotoMathPenRelativeX_A,
                (int)glyphs[1].XOffset,
                $"XOffset must be pen-relative. {NotoMathBaseRelativeX_A} means it is still "
                + "relative to the base glyph's origin.");

            Assert.AreEqual(NotoMathYOffset_A, (int)glyphs[1].YOffset, "YOffset must be unchanged");
        }

        [TestMethod]
        public void MarkOffset_OverLowercase_IsPenRelative_NotoSansMath()
        {
            var glyphs = ShapeBaseAndMark(NotoSansMathFamily, 'a');

            Assert.AreEqual(NotoMathPenRelativeX_a, (int)glyphs[1].XOffset, "XOffset must be pen-relative");
            Assert.AreEqual(NotoMathYOffset_a, (int)glyphs[1].YOffset, "YOffset must be unchanged");
        }

        [TestMethod]
        public void MarkOffset_OverCapital_IsPenRelative_Mulish()
        {
            var glyphs = ShapeBaseAndMark(MulishFamily, 'A');

            Assert.AreEqual(MulishPenRelativeX_A, (int)glyphs[1].XOffset, "XOffset must be pen-relative");
            Assert.AreEqual(MulishYOffset_A, (int)glyphs[1].YOffset, "YOffset must be unchanged");
        }

        [TestMethod]
        public void MarkOffset_OverLowercase_IsPenRelative_Mulish()
        {
            var glyphs = ShapeBaseAndMark(MulishFamily, 'a');

            Assert.AreEqual(MulishPenRelativeX_a, (int)glyphs[1].XOffset, "XOffset must be pen-relative");
            Assert.AreEqual(MulishYOffset_a, (int)glyphs[1].YOffset, "YOffset must be unchanged");
        }

        #endregion

        #region Offsets accumulate

        [TestMethod]
        public void MarkOffset_DoesNotOverwriteExistingPlacement()
        {
            // SinglePos writes XPlacement/YPlacement into the same fields before mark positioning
            // runs (TextShaper.ApplyValueRecord uses +=). MarkToBaseProvider must accumulate, not
            // assign, or that contribution is silently discarded.
            var font = TestFolderEngine.LoadFont(NotoSansMathFamily, FontSubFamily.Regular);
            Assert.IsNotNull(font, $"{NotoSansMathFamily} must be present in the test font folder");

            const short seededX = 25;
            const short seededY = -40;

            var glyphs = BuildGlyphPair(font!, 'A', seededX, seededY);

            new MarkToBaseProvider(font!).ApplyMarkPositioning(glyphs, "latn", null);

            Assert.AreEqual(
                NotoMathPenRelativeX_A + seededX,
                (int)glyphs[1].XOffset,
                "an XOffset already present must be added to, not replaced");

            Assert.AreEqual(
                NotoMathYOffset_A + seededY,
                (int)glyphs[1].YOffset,
                "a YOffset already present must be added to, not replaced");
        }

        [TestMethod]
        public void MarkOffset_UsesKernedAdvance_NotRawAdvance()
        {
            // Mark positioning runs after kerning, so the base glyph's XAdvance may already have
            // been reduced. The pen is where XAdvance says it is, so that is the value the offset
            // must be taken from - not the raw hmtx advance.
            var font = TestFolderEngine.LoadFont(NotoSansMathFamily, FontSubFamily.Regular);
            Assert.IsNotNull(font);

            var glyphs = BuildGlyphPair(font!, 'A', 0, 0);

            // Simulate a kern pair having tightened the base advance by 30 font units.
            const short kerning = -30;
            glyphs[0].XAdvance = (short)(glyphs[0].XAdvance + kerning);

            new MarkToBaseProvider(font!).ApplyMarkPositioning(glyphs, "latn", null);

            Assert.AreEqual(
                NotoMathPenRelativeX_A - kerning,
                (int)glyphs[1].XOffset,
                "the offset must follow the kerned advance, since that is where the pen ends up");
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Shapes a base character followed by a combining acute and returns the two glyphs.
        /// </summary>
        private static List<ShapedGlyph> ShapeBaseAndMark(string family, char baseChar)
        {
            var font = TestFolderEngine.LoadFont(family, FontSubFamily.Regular);
            Assert.IsNotNull(font, $"{family} must be present in the test font folder");

            var shaped = new TextShaper(TestFolderEngine, font!).Shape($"{baseChar}{CombiningAcute}");

            Assert.AreEqual(
                2,
                shaped.Glyphs.Length,
                "the sequence must stay decomposed - a ccmp/liga substitution would defeat the test");

            return new List<ShapedGlyph>(shaped.Glyphs);
        }

        /// <summary>
        /// Builds a base glyph plus combining acute directly, so the test controls XAdvance and any
        /// offset already present. Advances come from hmtx, matching what TextShaper would set.
        /// </summary>
        private static List<ShapedGlyph> BuildGlyphPair(OpenTypeFont font, char baseChar, short seededX, short seededY)
        {
            ushort baseGlyphId = font.CmapTable.GetGlyphId(baseChar);
            ushort markGlyphId = font.CmapTable.GetGlyphId(CombiningAcute);

            Assert.AreNotEqual(0, (int)baseGlyphId, $"'{baseChar}' must be in the font");
            Assert.AreNotEqual(0, (int)markGlyphId, "the combining acute must be in the font");

            short baseAdvance = (short)font.HmtxTable.GetAdvanceWidth(baseGlyphId);
            short markAdvance = (short)font.HmtxTable.GetAdvanceWidth(markGlyphId);

            return new List<ShapedGlyph>
            {
                new ShapedGlyph
                {
                    GlyphId = baseGlyphId,
                    XAdvance = baseAdvance,
                    BaseAdvance = baseAdvance
                },
                new ShapedGlyph
                {
                    GlyphId = markGlyphId,
                    XAdvance = markAdvance,
                    BaseAdvance = markAdvance,
                    XOffset = seededX,
                    YOffset = seededY
                }
            };
        }

        #endregion
    }
}