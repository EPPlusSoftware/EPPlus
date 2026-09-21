/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           Unicode Variation Sequence subsetting parity test
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Parity test analogous to the measurement/render width-parity checks (see WP1's kerning
    /// fix): verifies that shaping against the FULL font (the measurement path,
    /// GetCellCollectionFromRange) and shaping against the ROUND-TRIPPED SUBSET font (the render
    /// path - what actually ends up embedded in the PDF) agree on a Unicode Variation Sequence.
    ///
    /// Glyph IDs are renumbered by subsetting, so this can't compare raw glyph IDs across the two
    /// fonts. Instead it compares BEHAVIOR: does each font's own shaping still tell the
    /// (base, selector) pair apart from the plain base character on its own terms?
    ///
    /// "Round-tripped" means serialized to bytes and reloaded - exactly like a font that has gone
    /// through PDF embedding. This exercises CmapSubsetProcessor's discovery/rewrite of the
    /// variation-sequence data AND CmapTable's serialization of the format-14 subtable, both of
    /// which are currently missing.
    /// </summary>
    [TestClass]
    public class VariationSequenceSubsettingParityTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const uint Vs01 = 0xFE00; // VARIATION SELECTOR-1 (BMP)

        /// <summary>
        /// Loads a private (non-cached) Roboto instance and registers 'A' + Vs01 as a variation
        /// sequence resolving to whatever glyph 'B' already maps to (a distinct, real, existing
        /// glyph - not a made-up one).
        /// </summary>
        private OpenTypeFont LoadFullFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB)
        {
            var font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular, true);

            Assert.IsTrue(font.CmapTable.TryGetGlyphId('A', out glyphIdA), "Test font is expected to contain 'A'.");
            Assert.IsTrue(font.CmapTable.TryGetGlyphId('B', out glyphIdB), "Test font is expected to contain 'B'.");
            Assert.AreNotEqual(glyphIdA, glyphIdB, "Test relies on 'A' and 'B' mapping to different glyphs.");

            var subtable14 = new CmapSubtable14();
            subtable14.VariationSelectors.Add(new VariationSelector
            {
                VarSelector = Vs01,
                NonDefaultUvsTable = new NonDefaultUvsTable
                {
                    Mappings = new List<UvsMapping>
                    {
                        new UvsMapping { UnicodeValue = 'A', GlyphId = glyphIdB }
                    }
                }
            });
            font.CmapTable.SubTables.Add(subtable14);

            return font;
        }

        [TestMethod]
        public void Shape_RoundTrippedSubsetFont_StillDistinguishesVariationSequenceFromPlainBaseChar()
        {
            const string textWithSequence = "A\uFE00";

            var fullFont = LoadFullFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB);

            // --- Measurement path: shape against the full font (GetCellCollectionFromRange today
            // shapes here) ---
            var fullShaper = new TextShaper(TestFolderEngine, fullFont);
            var fullPlain = fullShaper.Shape("A");
            var fullSequence = fullShaper.Shape(textWithSequence);

            Assert.AreEqual(1, fullSequence.Glyphs.Length);
            Assert.AreEqual(2, fullSequence.Glyphs[0].CharCount, "Sanity check: the full font must consume the pair as one glyph.");
            Assert.AreNotEqual(fullPlain.Glyphs[0].GlyphId, fullSequence.Glyphs[0].GlyphId,
                "Sanity check: the variation sequence must select a different glyph than plain 'A' in the full font.");

            // --- Render path: subset for exactly this text, then round-trip (serialize + reload) -
            // exactly what happens when the subset is embedded in, and later read back from, a PDF. ---
            var subsetFont = fullFont.CreateSubset(textWithSequence);
            var subsetBytes = subsetFont.Serialize();
            var roundTrippedSubsetFont = new OpenTypeFont(subsetBytes);

            var subsetShaper = new TextShaper(TestFolderEngine, roundTrippedSubsetFont);
            var subsetPlain = subsetShaper.Shape("A");
            var subsetSequence = subsetShaper.Shape(textWithSequence);

            Assert.AreEqual(1, subsetSequence.Glyphs.Length);
            Assert.AreEqual(2, subsetSequence.Glyphs[0].CharCount,
                "The round-tripped SUBSET font must still consume the pair as one glyph. Today it silently " +
                "falls back to the base glyph, because CmapSubsetProcessor doesn't discover/preserve format-14 " +
                "data for the subset, and CmapTable.Serialize unconditionally drops format-14 subtables.");
            Assert.AreNotEqual(subsetPlain.Glyphs[0].GlyphId, subsetSequence.Glyphs[0].GlyphId,
                "The render path (round-tripped subset) must select a DIFFERENT glyph for the variation " +
                "sequence than for plain 'A' - matching what the measurement path (full font) does above.");
        }
    }
}