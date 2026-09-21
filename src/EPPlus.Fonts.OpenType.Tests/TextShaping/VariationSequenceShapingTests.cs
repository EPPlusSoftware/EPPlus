/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           Unicode Variation Sequence shaping tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Verifies that TextShaper actually consults cmap format 14 (Unicode Variation Sequences)
    /// during shaping - today the format-14 subtable is parsed but MapToGlyphs never looks at it.
    ///
    /// The variation-sequence data is injected synthetically into a real, otherwise-unmodified
    /// test font, so the tests don't depend on any specific font shipping real UVS data.
    ///
    /// IMPORTANT: TestFolderEngine.LoadFont caches and freezes (IsReadOnly) the fonts it returns,
    /// and that cache is shared across the whole test run. These tests load with ignoreCache: true
    /// so the synthetic format-14 subtable is only ever added to a private, per-test font instance.
    /// </summary>
    [TestClass]
    public class VariationSequenceShapingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private const uint Vs01 = 0xFE00;               // VARIATION SELECTOR-1 (BMP, 1 UTF-16 char)
        private const uint Vs17Supplementary = 0xE0100;  // VARIATION SELECTOR-17 (supplementary plane, surrogate pair)

        /// <summary>
        /// Loads a private (non-cached) instance of Roboto and adds a synthetic format-14
        /// subtable registering 'A' + Vs01 and 'A' + Vs17Supplementary as variation sequences
        /// that both resolve to the glyph 'B' already maps to (a visibly different, existing
        /// glyph). 'B' itself is left completely untouched.
        /// </summary>
        private OpenTypeFont LoadFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB)
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

            subtable14.VariationSelectors.Add(new VariationSelector
            {
                VarSelector = Vs17Supplementary,
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
        public void Shape_BaseCharPlusBmpVariationSelector_ConsumesBothIntoOneVariantGlyph()
        {
            var font = LoadFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("A\uFE00");

            Assert.AreEqual(1, shaped.Glyphs.Length, "The base char and the variation selector should collapse into a single glyph.");
            Assert.AreEqual(glyphIdB, shaped.Glyphs[0].GlyphId, "Should resolve to the variant glyph registered in cmap format 14, not glyphIdA.");
            Assert.AreEqual(2, shaped.Glyphs[0].CharCount, "Should consume both UTF-16 chars (1-char base + 1-char selector).");
            Assert.AreEqual(0, shaped.Glyphs[0].ClusterIndex);
        }

        [TestMethod]
        public void Shape_BaseCharPlusSupplementaryPlaneVariationSelector_ConsumesAllThreeChars()
        {
            var font = LoadFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            string selector = char.ConvertFromUtf32((int)Vs17Supplementary); // 2 UTF-16 chars (surrogate pair)
            var shaped = shaper.Shape("A" + selector);

            Assert.AreEqual(1, shaped.Glyphs.Length, "The base char and the surrogate-pair selector should collapse into a single glyph.");
            Assert.AreEqual(glyphIdB, shaped.Glyphs[0].GlyphId);
            Assert.AreEqual(3, shaped.Glyphs[0].CharCount, "Should consume base (1 char) + selector (2 chars, surrogate pair).");
        }

        [TestMethod]
        public void Shape_UnregisteredBaseCharPlusVariationSelector_DoesNotConsumeSelector()
        {
            var font = LoadFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            // 'B' has no entry under Vs01 in our synthetic table, so this sequence isn't registered.
            var shaped = shaper.Shape("B\uFE00");

            Assert.AreEqual(2, shaped.Glyphs.Length, "An unregistered sequence must fall back to two separate glyphs, not be silently consumed.");
            Assert.AreEqual(glyphIdB, shaped.Glyphs[0].GlyphId, "'B' should still resolve to its own ordinary glyph.");
            Assert.AreEqual(1, shaped.Glyphs[0].CharCount);
        }

        [TestMethod]
        public void Shape_PlainCharWithoutSelector_IsUnaffected()
        {
            var font = LoadFontWithSyntheticVariationSequence(out ushort glyphIdA, out ushort glyphIdB);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("A");

            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreEqual(glyphIdA, shaped.Glyphs[0].GlyphId, "Without a following selector, 'A' must still map to its own glyph.");
            Assert.AreEqual(1, shaped.Glyphs[0].CharCount);
        }
    }
}