/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Kerning from extension wrapped GPOS lookups
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.TextShaping.Kerning;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    /// <summary>
    /// Tests that pair positioning is actually FOUND for a given font, both in the full font and
    /// in a serialized subset.
    ///
    /// These tests exist because every bug they cover was silent. GposKerningProvider required
    /// LookupType == 2 and ExtensionPosHandler only handled MarkToBase, so a font whose kern
    /// lookups are extension wrapped simply came out unkerned with no error anywhere. Asserting
    /// that shaping does not throw is not enough - the assertion has to be that a specific
    /// adjustment for a specific pair is retrieved.
    /// </summary>
    [TestClass]
    public class GposExtensionKerningTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        /// <summary>
        /// EB Garamond has its kern feature as lookup type 9 (Extension) wrapping PairPos type 2,
        /// with both format 1 and format 2 subtables. This is the same failure mode as Calibri,
        /// without needing a system font. Units per em is 1000.
        /// </summary>
        private const string GaramondFamily = "EB Garamond";

        /// <summary>
        /// Roboto has its kern feature as PairPos type 2 directly, and 2048 units per em.
        /// Used as the guard that unwrapped lookups still work, and as the font that exposes
        /// units per em scaling errors as a factor of roughly two.
        /// </summary>
        private const string RobotoFamily = "Roboto";

        // Adjustments read from the kern feature of the shipped test fonts, in font units.
        private const int GaramondKerningAV = -140;   // class based (format 2) subtable
        private const int GaramondKerningTo = -105;   // class based (format 2) subtable
        private const int RobotoKerningAV = -87;      // class based (format 2) subtable
        private const int RobotoKerningFa = -34;      // individual pair (format 1) subtable

        #region Full font

        [TestMethod]
        public void FullFont_ExtensionWrappedKernLookup_IsFoundByKerningProvider()
        {
            var font = TestFolderEngine.LoadFont(GaramondFamily, FontSubFamily.Regular);
            Assert.IsNotNull(font, $"{GaramondFamily} must be present in the test font folder");
            Assert.IsNotNull(font.GposTable, $"{GaramondFamily} must have a GPOS table");

            var provider = new GposKerningProvider(font.GposTable);

            Assert.AreEqual(
                GaramondKerningAV,
                (int)provider.GetKerning(font.CmapTable.GetGlyphId('A'), font.CmapTable.GetGlyphId('V')),
                "A+V kerning must be found even though the kern lookup is extension wrapped");

            Assert.AreEqual(
                GaramondKerningTo,
                (int)provider.GetKerning(font.CmapTable.GetGlyphId('T'), font.CmapTable.GetGlyphId('o')),
                "T+o kerning must be found even though the kern lookup is extension wrapped");
        }

        [TestMethod]
        public void FullFont_DirectKernLookup_IsStillFoundByKerningProvider()
        {
            // Guards the removal of the LookupType == 2 filter in GposKerningProvider:
            // lookups that are not extension wrapped must keep working.
            var font = TestFolderEngine.LoadFont(RobotoFamily, FontSubFamily.Regular);
            Assert.IsNotNull(font.GposTable);

            var provider = new GposKerningProvider(font.GposTable);

            Assert.AreEqual(
                RobotoKerningAV,
                (int)provider.GetKerning(font.CmapTable.GetGlyphId('A'), font.CmapTable.GetGlyphId('V')),
                "A+V comes from a class based subtable");

            Assert.AreEqual(
                RobotoKerningFa,
                (int)provider.GetKerning(font.CmapTable.GetGlyphId('F'), font.CmapTable.GetGlyphId('a')),
                "F+a comes from an individual pair subtable");
        }

        [TestMethod]
        public void Shape_ExtensionWrappedKerning_IsAppliedToLeftGlyphAdvance()
        {
            var font = TestFolderEngine.LoadFont(GaramondFamily, FontSubFamily.Regular);
            var shaper = new TextShaper(TestFolderEngine, font);

            var shaped = shaper.Shape("AV");

            Assert.AreEqual(2, shaped.Glyphs.Length);

            // TextShaper.ApplyKerning adjusts the advance of the LEFT glyph of the pair.
            int kerning = shaped.Glyphs[0].XAdvance - shaped.Glyphs[0].BaseAdvance;
            Assert.AreEqual(GaramondKerningAV, kerning, "kerning must reach the shaped advance");

            Assert.AreEqual(
                0,
                shaped.Glyphs[1].XAdvance - shaped.Glyphs[1].BaseAdvance,
                "the right glyph of the pair must not be adjusted");
        }

        #endregion

        #region Subset

        [TestMethod]
        public void Subset_ExtensionWrappedKernLookup_SurvivesWithInnerLookupType()
        {
            var subset = FontTestHelper.RoundtripSubset(TestFolderEngine, GaramondFamily, "AVATAR Wa To");

            Assert.IsNotNull(subset.GposTable, "GPOS must survive subsetting");

            var kernLookups = GetKernLookups(subset.GposTable);
            Assert.AreNotEqual(0, kernLookups.Count, "the kern feature must still reference lookups");

            int pairPosSubtables = 0;
            foreach (var lookup in kernLookups)
            {
                foreach (var subtable in lookup.SubTables)
                {
                    if (subtable is PairPosSubTable)
                    {
                        pairPosSubtables++;

                        // The subset must carry the inner lookup type, not the extension type.
                        // Re-wrapping as type 9 without a real ExtensionPosFormat1 record is what
                        // made the embedded GPOS table unreadable to external parsers.
                        Assert.AreEqual(
                            2,
                            (int)lookup.LookupType,
                            "a lookup holding PairPos subtables must be declared as type 2");
                    }
                }
            }

            Assert.AreNotEqual(
                0,
                pairPosSubtables,
                "extension wrapped PairPos must not be dropped during subsetting");
        }

        [TestMethod]
        public void Subset_ExtensionWrappedKerning_IsStillFoundAfterRoundtrip()
        {
            var subset = FontTestHelper.RoundtripSubset(TestFolderEngine, GaramondFamily, "AVATAR Wa To");

            var provider = new GposKerningProvider(subset.GposTable);

            // Glyph ids are the subset ones here, which is exactly what the PDF path shapes with.
            Assert.AreEqual(
                GaramondKerningAV,
                (int)provider.GetKerning(subset.CmapTable.GetGlyphId('A'), subset.CmapTable.GetGlyphId('V')),
                "A+V kerning must survive subsetting and glyph id remapping");

            Assert.AreEqual(
                GaramondKerningTo,
                (int)provider.GetKerning(subset.CmapTable.GetGlyphId('T'), subset.CmapTable.GetGlyphId('o')),
                "T+o kerning must survive subsetting and glyph id remapping");
        }

        [TestMethod]
        public void Subset_ExpandedPairSets_AreSortedBySecondGlyph()
        {
            // PairPosHandler.RewriteFormat2 expands the class matrix by iterating
            // FontSubsettingContext.IncludedGlyphs, which is a HashSet with no defined
            // enumeration order. Unsorted records break the binary search in
            // PairPosSubTableFormat1.TryGetPairAdjustment and silently lose pairs.
            var subset = FontTestHelper.RoundtripSubset(
                TestFolderEngine, GaramondFamily, "AVATAR Wa To vwxyz .,-");

            int checkedPairSets = 0;

            foreach (var subtable in GetPairPosFormat1Subtables(subset.GposTable))
            {
                foreach (var pairSet in subtable.PairSets)
                {
                    if (pairSet?.PairValueRecords == null)
                    {
                        continue;
                    }

                    checkedPairSets++;

                    for (int i = 1; i < pairSet.PairValueRecords.Count; i++)
                    {
                        Assert.IsTrue(
                            pairSet.PairValueRecords[i - 1].SecondGlyph < pairSet.PairValueRecords[i].SecondGlyph,
                            "PairValueRecords must be strictly ascending by SecondGlyph, but "
                            + $"{pairSet.PairValueRecords[i - 1].SecondGlyph} came before "
                            + $"{pairSet.PairValueRecords[i].SecondGlyph}");
                    }
                }
            }

            Assert.AreNotEqual(0, checkedPairSets, "the test must actually have pair sets to check");
        }

        [TestMethod]
        public void Subset_EveryPairInExpandedPairSet_IsFoundByBinarySearch()
        {
            // The observable symptom of unsorted records: a pair is present in the table but
            // cannot be retrieved.
            var subset = FontTestHelper.RoundtripSubset(
                TestFolderEngine, GaramondFamily, "AVATAR Wa To vwxyz .,-");

            int checkedPairs = 0;

            foreach (var subtable in GetPairPosFormat1Subtables(subset.GposTable))
            {
                var coverage = subtable.Coverage as CoverageTableFormat1;
                Assert.IsNotNull(coverage, "subsetted coverage is always written as format 1");

                for (int i = 0; i < coverage!.GlyphArray.Length && i < subtable.PairSets.Count; i++)
                {
                    var firstGlyph = coverage.GlyphArray[i];
                    var pairSet = subtable.PairSets[i];

                    if (pairSet?.PairValueRecords == null)
                    {
                        continue;
                    }

                    foreach (var record in pairSet.PairValueRecords)
                    {
                        checkedPairs++;

                        Assert.IsTrue(
                            subtable.TryGetPairAdjustment(firstGlyph, record.SecondGlyph, out _, out _),
                            $"pair {firstGlyph}+{record.SecondGlyph} is in the table but was not found");
                    }
                }
            }

            Assert.AreNotEqual(0, checkedPairs, "the test must actually have pairs to check");
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Returns the lookups referenced by the kern feature, resolved through the feature list.
        /// </summary>
        private static List<LookupTable> GetKernLookups(GposTable gpos)
        {
            var result = new List<LookupTable>();

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value != "kern")
                {
                    continue;
                }

                foreach (var lookupIndex in featureRecord.FeatureTable.LookupListIndices)
                {
                    if (lookupIndex < gpos.LookupList.Lookups.Count)
                    {
                        result.Add(gpos.LookupList.Lookups[lookupIndex]);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Returns every PairPos format 1 subtable in the kern feature. After subsetting both
        /// original formats are written as format 1, since RewriteFormat2 expands the class matrix.
        /// </summary>
        private static List<PairPosSubTableFormat1> GetPairPosFormat1Subtables(GposTable gpos)
        {
            var result = new List<PairPosSubTableFormat1>();

            foreach (var lookup in GetKernLookups(gpos))
            {
                foreach (var subtable in lookup.SubTables)
                {
                    if (subtable is PairPosSubTableFormat1 format1 && format1.PairSets != null)
                    {
                        result.Add(format1);
                    }
                }
            }

            return result;
        }

        #endregion
    }
}