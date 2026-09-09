/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           Unicode Variation Sequence lookup tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.Tables.Cmap
{
    /// <summary>
    /// Tests the new CmapTable.TryGetGlyphId(uint baseCodePoint, uint variationSelector, out ushort glyphId)
    /// overload, which looks a (base char, variation selector) pair up against the format-14
    /// subtable that is already parsed but currently never consulted.
    ///
    /// All tables here are built by hand so the tests are deterministic and don't depend on any
    /// specific font shipping real Unicode Variation Sequence data.
    /// </summary>
    [TestClass]
    public class CmapVariationSequenceTests
    {
        public TestContext? TestContext { get; set; }

        // Variation selectors (see Unicode Standard Annex #38 / ISO 10646).
        private const uint Vs01 = 0xFE00;          // VARIATION SELECTOR-1 (BMP)
        private const uint Vs17 = 0xE0100;         // VARIATION SELECTOR-17 (supplementary plane)
        private const uint UnknownSelector = 0xFE0F; // VARIATION SELECTOR-16 - not registered in our synthetic table

        // Arbitrary CJK base characters used only as stand-ins in the synthetic table.
        private const uint BaseCharNonDefault = 0x4E00; // 一
        private const uint BaseCharDefault = 0x4E01;    // 丁
        private const uint BaseCharUnmapped = 0x4E02;   // 丂 - not registered under any selector

        private const ushort NonDefaultGlyphId = 500;
        private const ushort DefaultCmapGlyphId = 42;

        /// <summary>
        /// Builds a CmapTable with an ordinary Format 4 subtable (for the "default UVS" fallback
        /// to have something to fall back to) plus a Format 14 subtable registering:
        ///   - BaseCharNonDefault + Vs01  -> explicit override glyph (NonDefaultUvsTable)
        ///   - BaseCharDefault    + Vs17  -> "use the base character's ordinary glyph" (DefaultUvsTable)
        /// </summary>
        private static CmapTable BuildSyntheticCmapTable()
        {
            var cmap = new CmapTable();

            var baseMapping = new Dictionary<uint, ushort> { { BaseCharDefault, DefaultCmapGlyphId } };
            cmap.SubTables.Add(CmapFormat4.CreateFromMappings(baseMapping));

            var subtable14 = new CmapSubtable14();

            var nonDefaultSelector = new VariationSelector
            {
                VarSelector = Vs01,
                NonDefaultUvsTable = new NonDefaultUvsTable
                {
                    Mappings = new List<UvsMapping>
                    {
                        new UvsMapping { UnicodeValue = BaseCharNonDefault, GlyphId = NonDefaultGlyphId }
                    }
                }
            };
            subtable14.VariationSelectors.Add(nonDefaultSelector);

            var defaultSelector = new VariationSelector
            {
                VarSelector = Vs17,
                DefaultUvsTable = new DefaultUvsTable
                {
                    Ranges = new List<UnicodeRange>
                    {
                        new UnicodeRange { StartUnicodeValue = BaseCharDefault, AdditionalCount = 0 }
                    }
                }
            };
            subtable14.VariationSelectors.Add(defaultSelector);

            cmap.SubTables.Add(subtable14);

            return cmap;
        }

        [TestMethod]
        public void TryGetGlyphId_NonDefaultUvs_ReturnsExplicitOverrideGlyph()
        {
            var cmap = BuildSyntheticCmapTable();

            bool found = cmap.TryGetGlyphId(BaseCharNonDefault, Vs01, out ushort glyphId);

            Assert.IsTrue(found, "A registered non-default variation sequence should resolve.");
            Assert.AreEqual(NonDefaultGlyphId, glyphId);
        }

        [TestMethod]
        public void TryGetGlyphId_DefaultUvs_FallsBackToOrdinaryCmapGlyph()
        {
            var cmap = BuildSyntheticCmapTable();

            bool found = cmap.TryGetGlyphId(BaseCharDefault, Vs17, out ushort glyphId);

            Assert.IsTrue(found, "A registered default variation sequence should still resolve (it's a valid, known sequence).");
            Assert.AreEqual(DefaultCmapGlyphId, glyphId,
                "Default UVS entries don't carry their own glyph id - they mean 'use the base character's ordinary cmap glyph'.");
        }

        [TestMethod]
        public void TryGetGlyphId_SelectorRegisteredButBaseCharIsNot_ReturnsFalse()
        {
            var cmap = BuildSyntheticCmapTable();

            bool found = cmap.TryGetGlyphId(BaseCharUnmapped, Vs01, out ushort glyphId);

            Assert.IsFalse(found, "A base character absent from both the default and non-default UVS tables is not a registered sequence.");
            Assert.AreEqual(0, glyphId);
        }

        [TestMethod]
        public void TryGetGlyphId_UnknownVariationSelector_ReturnsFalse()
        {
            var cmap = BuildSyntheticCmapTable();

            bool found = cmap.TryGetGlyphId(BaseCharNonDefault, UnknownSelector, out ushort glyphId);

            Assert.IsFalse(found, "A variation selector with no entry at all in the format-14 subtable is not a registered sequence.");
        }

        [TestMethod]
        public void TryGetGlyphId_NoFormat14Subtable_ReturnsFalse()
        {
            var cmap = new CmapTable();
            cmap.SubTables.Add(CmapFormat4.CreateFromMappings(new Dictionary<uint, ushort> { { BaseCharDefault, DefaultCmapGlyphId } }));

            bool found = cmap.TryGetGlyphId(BaseCharDefault, Vs01, out ushort glyphId);

            Assert.IsFalse(found, "Without a format-14 subtable there is no Unicode Variation Sequence data to consult at all.");
        }
    }
}