/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/22/2025         EPPlus Software AB           Regression tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Runtime.InteropServices;

namespace EPPlus.Fonts.OpenType.Tests.Regression
{
    /// <summary>
    /// Tests for specific bugs that have been found and fixed.
    /// Each test documents a bug to prevent regression.
    /// </summary>
    [TestClass]
    public class RegressionTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void Bug_20251222_CircularLigatureDependency_Roboto()
        {
            // BUG: ffi ligature referenced fi ligature as component (GID >= 400)
            //      This caused circular dependencies in subsetting
            // 
            // Original structure in Roboto:
            //   fi-ligature (444): components = [i(77), mysterious-447]
            //   ffi-ligature (446): components = [f(74), i(77), fi-ligature(444)]
            //   ff-ligature (443): components = [f(74), 449]
            //
            // CAUSE: Discovery phase included ligatures as components, creating cycles
            // FIX: Ignore components with GID >= 400 (ligature glyphs) in both
            //      Discovery and Rewrite phases
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("ffi");

            SaveFont("regression_ffi_circular.ttf", subset);

            // Should create valid font with ffi ligature
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            Assert.IsNotNull(parsed.GsubTable);

            int ligCount = FontTestHelper.CountLigatures(parsed);
            Assert.IsTrue(ligCount > 0, "ffi ligature should exist after subsetting");

            FontTestHelper.AssertFontValid(parsed);
        }

        [TestMethod]
        public void Bug_20251222_AbcSubset_TooManyGlyphs()
        {
            // BUG: Subsetting 'abc' resulted in 28 glyphs instead of ~10
            // CAUSE: Ligature discovery included all f-ligatures (fi, ff, fl, ffi, ffl)
            //        even though 'abc' doesn't contain 'f'
            // ROOT CAUSE: Discovery didn't validate that base character components existed
            // FIX: Only include ligatures where ALL base character components (GID < 400)
            //      exist in the subset
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset(new[] { 'a', 'b', 'c' });

            SaveFont("regression_abc_glyphs.ttf", subset);

            // Should have ~10 glyphs (abc + space + .notdef + variants), NOT 28
            Assert.IsTrue(subset.MaxpTable.numGlyphs >= 5 && subset.MaxpTable.numGlyphs <= 15,
                $"Expected 5-15 glyphs for abc, got {subset.MaxpTable.numGlyphs}");

            // Should have NO ligatures
            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.AreEqual(0, ligCount, "abc should have no ligatures");
        }

        [TestMethod]
        public void Bug_20251222_Fiffig_LostLigatures_AfterAbcFix()
        {
            // BUG: After fixing 'abc' bug, 'fiffig' lost all ligatures
            // CAUSE: Fix was too aggressive - also filtered out valid ligatures
            // FIX: Corrected logic to only check base components (GID < 400)
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fiffig");

            SaveFont("regression_fiffig_ligatures.ttf", subset);

            // Should have exactly 3 ligatures: fi, ff, ffi
            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.AreEqual(3, ligCount, "fiffig should have fi, ff, ffi ligatures");

            // Verify font is valid
            FontTestHelper.AssertFontValid(subset);
        }

        [TestMethod]
        public void Bug_20251222_FeaturePointsToInvalidLookup()
        {
            // BUG: After subsetting, features pointed to non-existent lookup indices
            // CAUSE: Lookups were removed during subsetting but feature indices weren't remapped
            // FIX: FeatureListTable.Rewrite now uses lookupMap to remap indices
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fiffig");

            SaveFont("regression_feature_lookup.ttf", subset);

            // Verify all features point to valid lookups
            if (subset.GsubTable?.FeatureList?.FeatureRecords != null)
            {
                int lookupCount = subset.GsubTable.LookupList?.Lookups?.Count ?? 0;

                foreach (var feature in subset.GsubTable.FeatureList.FeatureRecords)
                {
                    if (feature.FeatureTable?.LookupListIndices != null)
                    {
                        foreach (var idx in feature.FeatureTable.LookupListIndices)
                        {
                            Assert.IsTrue(idx < lookupCount,
                                $"Feature {feature.FeatureTag} points to invalid lookup {idx} (max: {lookupCount - 1})");
                        }
                    }
                }
            }
        }

        [TestMethod]
        public void Bug_20251222_LigatureComponentRewrite_WrongDictionary()
        {
            // BUG: LigatureSetTable.Rewrite used wrong dictionary for component mapping
            // CAUSE: Used oldLig.Components instead of looking up in OldToNewGlyphId
            // SYMPTOM: Components weren't remapped to new GIDs
            // FIX: Properly lookup each component in OldToNewGlyphId dictionary
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fi");

            SaveFont("regression_ligature_components.ttf", subset);

            // Serialize and re-parse to verify components are correct
            var bytes = subset.Serialize();
            var parsed = new OpenTypeFont(bytes, font.Format);

            // Should have fi ligature with correctly remapped components
            int ligCount = FontTestHelper.CountLigatures(parsed);
            Assert.IsTrue(ligCount >= 1, "Should have fi ligature");

            FontTestHelper.AssertFontValid(parsed);
        }

        [TestMethod]
        public void Bug_20251222_ValidationCrash_NullRawData()
        {
            // BUG: TableRecordsValidator crashed with NullReferenceException on subset fonts
            // CAUSE: Validator tried to access font.RawData which is null for in-memory fonts
            // LINE: TableRecordsValidator.cs line 111: byte[] fontData = (byte[])font.RawData.Clone()
            // FIX: Check if RawData is null before accessing, skip checksum validation for in-memory fonts
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("f");

            SaveFont("regression_validation_rawdata.ttf", subset);

            // Should not crash during validation
            FontTestHelper.AssertFontValid(subset);

            Assert.IsNotNull(subset);
            Assert.IsTrue(subset.MaxpTable.numGlyphs > 0);
        }

        [TestMethod]
        public void Bug_20251222_ValidationCrash_FileLengthZero()
        {
            // BUG: TableRecordsValidator failed with "Font file length could not be determined or is zero"
            // CAUSE: font.FileLength returned 0 for in-memory fonts (no underlying stream)
            // FIX: Calculate fileLength from TableRecords if FileLength property is <= 0
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fl");

            SaveFont("regression_validation_filelength.ttf", subset);

            // Should not fail with file length error
            FontTestHelper.AssertFontValid(subset);

            Assert.IsNotNull(subset);
        }

        [TestMethod]
        public void Bug_20251222_MissingCoverageInitialization()
        {
            // BUG: Coverage table and SubtableFormat not initialized in ligature rewrite
            // CAUSE: Missing initialization in LigatureSubstSubTable.Rewrite
            // SYMPTOM: Serialization failed or produced invalid fonts
            // FIX: Initialize Coverage.SubstFormat = 1 and newSubTable.SubstFormat = 1
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("office");

            SaveFont("regression_coverage_init.ttf", subset);

            // Should serialize successfully
            var bytes = subset.Serialize();
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length > 0);

            // Should parse successfully
            var parsed = new OpenTypeFont(bytes, font.Format);
            Assert.IsNotNull(parsed);

            FontTestHelper.AssertFontValid(parsed);
        }

        [TestMethod]
        public void Bug_20251222_EmptySubset_ShouldThrow()
        {
            // BUG: CreateSubset with empty string didn't throw exception
            // EXPECTED: ArgumentException for empty input
            // FIX: Added validation to throw ArgumentException if usedChars is empty
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            Assert.ThrowsException<ArgumentException>(() => font.CreateSubset(""));
            Assert.ThrowsException<ArgumentNullException>(() => font.CreateSubset((char[])null));
        }

        [TestMethod]
        public void Bug_20251222_CompoundLigatureComponents()
        {
            // BUG: Roboto's ligature definitions used compound structures where
            //      ligatures referenced other ligatures as components
            // EXAMPLE:
            //   fi-ligature (444): [i, mysterious-447] where 447 is another ligature
            //   ffi-ligature (446): [f, i, fi-ligature(444)]
            // FIX: Ignore ligature components (GID >= 400) in both Discovery and Rewrite
            // RESULT: Simplified component lists (ffi = [f, f, i] instead of [f, i, fi])
            // DATE: 2025-12-22

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var subset = font.CreateSubset("fiffig");

            SaveFont("regression_compound_components.ttf", subset);

            // Should have 3 ligatures: ff, fi, ffi
            int ligCount = FontTestHelper.CountLigatures(subset);
            Assert.AreEqual(3, ligCount);

            // Verify in FontDrop or similar tool that ligatures render correctly
            FontTestHelper.AssertFontValid(subset);
        }
    }
}