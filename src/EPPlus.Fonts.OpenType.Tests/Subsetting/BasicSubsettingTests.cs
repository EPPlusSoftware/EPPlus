/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Basic subsetting tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.Diagnostics;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class BasicSubsettingTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void Subset_Abc_RoundtripValidation()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_abc.ttf", parsedFont);

            // Assert: Check table presence
            Assert.AreEqual(12, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("loca"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("glyf"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("cmap"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("post"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("OS/2"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("GSUB"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("GPOS"));

            // Validate all tables
            FontTestHelper.AssertFontValid(parsedFont);

            // ✅ FIXED: abc should have NO ligatures
            int ligatureCount = FontTestHelper.CountLigatures(parsedFont);
            Assert.AreEqual(0, ligatureCount, "abc should have NO ligatures");

            // Verify glyph count (approximately)
            int expectedGlyphs = 3;      // a, b, c
            expectedGlyphs += 1;         // + space (U+0020)
            expectedGlyphs += 1;         // + .notdef (GID 0)
            expectedGlyphs += 5;         // + variants from Single Substitution

            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.HheaTable.numberOfHMetrics);

            // Verify space exists in cmap
            Assert.IsTrue(parsedFont.CmapTable.ContainsChar(32));
        }

        [TestMethod]
        public void Subset_Fiffig_WithFullValidation()
        {
            // Arrange
            var fontName = "Roboto";
            var font = OpenTypeFonts.GetFontData(FontFolders, fontName, FontSubFamily.Regular, true);
            var subsetFont = font.CreateSubset("fiffig");

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_fiffig.ttf", parsedFont);

            // Assert
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_SingleChar_ShouldWork()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Mulish", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a' });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            SaveFont("subset_Mulish_a.ttf", parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_MultipleChars_ShouldWork()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] {
                'F', 'l', 'y', 'g', 'a', 'n', 'd', 'e', 'b', 'ä', 'c', 'k', 's', 'i', 'r', 'ö', 'h', 'w', 'p', 'å',
                'm', 'j', 'u', 't', 'v', 'o', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0'
            });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            SaveFont("subset_Roboto_flygande_bäckasiner.ttf", parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_RoundtripHelper_ShouldWork()
        {
            // Using FontTestHelper.RoundtripSubset
            var parsedFont = FontTestHelper.RoundtripSubset("Roboto", "test", FontFolders);

            SaveFont("subset_Roboto_test_via_helper.ttf", parsedFont);

            Assert.IsNotNull(parsedFont);
            Assert.IsTrue(parsedFont.MaxpTable.numGlyphs > 0);
        }

        [TestMethod]
        public void Check_Original_Roboto_Ligatures()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Get glyph IDs in ORIGINAL
            font.CmapTable.TryGetGlyphId('f', out ushort fGlyph);
            font.CmapTable.TryGetGlyphId('i', out ushort iGlyph);
            font.CmapTable.TryGetGlyphId('o', out ushort oGlyph);
            font.CmapTable.TryGetGlyphId('g', out ushort gGlyph);

            Debug.WriteLine("=== ORIGINAL ROBOTO ===");
            Debug.WriteLine($"'f' = glyph {fGlyph}");
            Debug.WriteLine($"'i' = glyph {iGlyph}");
            Debug.WriteLine($"'o' = glyph {oGlyph}");
            Debug.WriteLine($"'g' = glyph {gGlyph}");

            Debug.WriteLine("\n=== ORIGINAL LIGATURES ===");
            var ligLookups = font.GsubTable.LookupList.Lookups.Where(l => l.LookupType == 4).ToList();

            foreach (var lookup in ligLookups)
            {
                foreach (var subtable in lookup.SubTables)
                {
                    var ligSubtable = subtable as LigatureSubstSubTable;
                    if (ligSubtable != null && ligSubtable.LigatureSets.ContainsKey(fGlyph))
                    {
                        var ligSet = ligSubtable.LigatureSets[fGlyph];
                        Debug.WriteLine($"First glyph {fGlyph} ('f') has {ligSet.Ligatures.Count} ligatures:");

                        foreach (var lig in ligSet.Ligatures)
                        {
                            var componentIds = string.Join(", ", lig.Components.Select(c => c.ToString()).ToArray());
                            Debug.WriteLine($"  Output: {lig.LigatureGlyph}, Components: [{componentIds}]");
                        }
                    }
                }
            }
        }

        [TestMethod]
        public void Subset_Ligatures_ShouldStillWork()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Check ORIGINAL first
            Debug.WriteLine("=== ORIGINAL ROBOTO ===");
            if (font.GsubTable != null)
            {
                var origLigLookups = font.GsubTable.LookupList.Lookups.Where(l => l.LookupType == 4).ToList();
                Debug.WriteLine($"Original has {origLigLookups.Count} ligature lookups");

                font.CmapTable.TryGetGlyphId('f', out ushort origF);
                Debug.WriteLine($"'f' = glyph {origF} in original");

                foreach (var lookup in origLigLookups)
                {
                    foreach (var subtable in lookup.SubTables)
                    {
                        var ligSubtable = subtable as LigatureSubstSubTable;
                        if (ligSubtable != null && ligSubtable.LigatureSets.ContainsKey(origF))
                        {
                            var ligSet = ligSubtable.LigatureSets[origF];
                            Debug.WriteLine($"  'f' has {ligSet.Ligatures.Count} ligatures in original");
                        }
                    }
                }
            }

            Debug.WriteLine("\n=== ORIGINAL FEATURES ===");
            for (int i = 0; i < font.GsubTable.FeatureList.FeatureRecords.Count && i < 10; i++)
            {
                var feat = font.GsubTable.FeatureList.FeatureRecords[i];
                Debug.WriteLine($"Feature[{i}]: '{feat.FeatureTag.Value}'");
            }

            // Create subset
            Debug.WriteLine("\n=== CREATING SUBSET ===");
            var subsetFont = font.CreateSubset("fiffigoffice");

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            SaveFont("subset_Roboto_ligatures_test.ttf", parsedFont);

            Debug.WriteLine("\n=== SUBSET FONT ===");
            Debug.WriteLine($"Total glyphs: {parsedFont.MaxpTable.numGlyphs}");

            parsedFont.CmapTable.TryGetGlyphId('f', out ushort subsetF);
            parsedFont.CmapTable.TryGetGlyphId('i', out ushort subsetI);
            Debug.WriteLine($"'f' = glyph {subsetF}");
            Debug.WriteLine($"'i' = glyph {subsetI}");

            // Check GSUB
            Assert.IsNotNull(parsedFont.GsubTable, "GSUB should be present");

            Debug.WriteLine($"\n=== GSUB FEATURES (DETAILED) ===");
            foreach (var feature in parsedFont.GsubTable.FeatureList.FeatureRecords)
            {
                int lookupCount = feature.FeatureTable?.LookupListIndices?.Length ?? 0;
                Debug.WriteLine($"Feature: '{feature.FeatureTag.Value}', Lookups: {lookupCount}");
            }

            var ligLookups = parsedFont.GsubTable.LookupList.Lookups
                .Where(l => l.LookupType == 4)
                .ToList();

            Debug.WriteLine($"\n=== LIGATURE LOOKUPS ===");
            Debug.WriteLine($"Found {ligLookups.Count} ligature lookups");

            Assert.IsTrue(ligLookups.Count > 0, "Should have ligature lookups");

            // Check what ligatures exist
            foreach (var lookup in ligLookups)
            {
                foreach (var subtable in lookup.SubTables)
                {
                    var ligSubtable = subtable as LigatureSubstSubTable;
                    if (ligSubtable != null)
                    {
                        Debug.WriteLine($"\nLigature subtable has {ligSubtable.LigatureSets.Count} sets");

                        if (ligSubtable.LigatureSets.ContainsKey(subsetF))
                        {
                            var ligSet = ligSubtable.LigatureSets[subsetF];
                            Debug.WriteLine($"'f' (glyph {subsetF}) has {ligSet.Ligatures.Count} ligatures:");

                            foreach (var lig in ligSet.Ligatures)
                            {
                                var componentIds = string.Join(", ", lig.Components.Select(c => c.ToString()).ToArray());
                                Debug.WriteLine($"  Output: {lig.LigatureGlyph}, Components: [{componentIds}]");
                            }
                        }
                        else
                        {
                            Debug.WriteLine($"'f' (glyph {subsetF}) NOT FOUND in LigatureSets!");
                            Debug.WriteLine($"Available first glyphs: {string.Join(", ", ligSubtable.LigatureSets.Keys.Select(k => k.ToString()).ToArray())}");
                        }
                    }
                }
            }

            Debug.WriteLine($"\n=== SCRIPTLIST & LANGSYS ===");
            if (parsedFont.GsubTable.ScriptList != null)
            {
                foreach (var scriptRecord in parsedFont.GsubTable.ScriptList.ScriptRecords)
                {
                    string scriptTag = scriptRecord.ScriptTag.Value;
                    Debug.WriteLine($"\nScript: '{scriptTag}'");

                    var scriptTable = scriptRecord.ScriptTable;

                    // Check DefaultLangSys
                    if (scriptTable.DefaultLangSys != null)
                    {
                        var defLang = scriptTable.DefaultLangSys;
                        Debug.WriteLine($"  DefaultLangSys:");
                        Debug.WriteLine($"    RequiredFeatureIndex: {defLang.RequiredFeatureIndex}");
                        Debug.WriteLine($"    FeatureIndices: [{string.Join(", ", defLang.FeatureIndices.Select(i => i.ToString()).ToArray())}]");

                        // Show what features these indices point to
                        foreach (var featIdx in defLang.FeatureIndices)
                        {
                            if (featIdx < parsedFont.GsubTable.FeatureList.FeatureRecords.Count)
                            {
                                var feat = parsedFont.GsubTable.FeatureList.FeatureRecords[featIdx];
                                Debug.WriteLine($"      Feature[{featIdx}]: '{feat.FeatureTag.Value}'");
                            }
                        }
                    }
                }
            }
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }



        [TestMethod]
        public void Subset_WithGposKerning_ShouldPreservePositioning()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Use characters that ACTUALLY have kerning in Roboto
            var chars = new[] { 'f', 'e', 'c', 'd', 'g', 'E', 'a', 'b', ' ' };

            // Check ORIGINAL font - test f-e pair
            bool foundF_Original = font.CmapTable.TryGetGlyphId('f', out ushort fGlyphOrig);
            bool foundE_Original = font.CmapTable.TryGetGlyphId('e', out ushort eGlyphOrig);

            Debug.WriteLine($"=== ORIGINAL FONT ===");
            Debug.WriteLine($"f = glyph {fGlyphOrig}");
            Debug.WriteLine($"e = glyph {eGlyphOrig}");

            Assert.IsTrue(foundF_Original && foundE_Original, "Should find f and e in original");

            // Verify kerning exists in original
            bool hasOriginalKerning = false;
            if (font.GposTable != null)
            {
                var kernLookupsOrig = font.GposTable.LookupList.Lookups.Where(l => l.LookupType == 2).ToList();
                foreach (var lookup in kernLookupsOrig)
                {
                    var subtable = lookup.SubTables.FirstOrDefault() as PairPosSubTableFormat1;
                    if (subtable != null && subtable.TryGetPairAdjustment(fGlyphOrig, eGlyphOrig, out var val1, out var val2))
                    {
                        Debug.WriteLine($"ORIGINAL: f-e kerning = {val1.XAdvance}");
                        hasOriginalKerning = true;
                        break;
                    }
                }
            }

            Assert.IsTrue(hasOriginalKerning, "f-e should have kerning in original font");

            // Create subset
            Debug.WriteLine($"\n=== CREATING SUBSET ===");
            var subsetFont = font.CreateSubset(chars);

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            SaveFont("subset_Roboto_with_gpos_kerning.ttf", parsedFont);

            // Check SUBSET font
            bool foundF = parsedFont.CmapTable.TryGetGlyphId('f', out ushort fGlyph);
            bool foundE = parsedFont.CmapTable.TryGetGlyphId('e', out ushort eGlyph);

            Debug.WriteLine($"\n=== SUBSET FONT ===");
            Debug.WriteLine($"f = glyph {fGlyph}");
            Debug.WriteLine($"e = glyph {eGlyph}");

            Assert.IsNotNull(parsedFont.GposTable, "GPOS table should be present in subset");

            var kernLookups = parsedFont.GposTable.LookupList.Lookups.Where(l => l.LookupType == 2).ToList();
            Assert.IsTrue(kernLookups.Count > 0, "Should have at least one kerning lookup");

            // Verify f-e kerning is preserved
            bool hasKerning = false;
            foreach (var lookup in kernLookups)
            {
                var subtable = lookup.SubTables.FirstOrDefault() as PairPosSubTableFormat1;
                if (subtable != null)
                {
                    if (subtable.TryGetPairAdjustment(fGlyph, eGlyph, out var val1, out var val2))
                    {
                        Debug.WriteLine($"SUBSET: f-e kerning = {val1.XAdvance}");
                        hasKerning = true;
                        Assert.AreEqual(-24, val1.XAdvance, "Kerning value should match original");
                        break;
                    }
                }
            }

            Assert.IsTrue(hasKerning, "f-e kerning pair should be preserved in subset");
        }

        [TestMethod]
        public void Subset_WithGposSingleAdjustment_ShouldPreserve()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Include various characters that might have single adjustments
            var chars = new[] {
        'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
        'A', 'B', 'C', 'D', 'E', ' '
    };

            var subsetFont = font.CreateSubset(chars);

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_with_gpos_singleadj.ttf", parsedFont);

            // Assert
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

            // Check if SinglePos lookups exist
            if (parsedFont.GposTable != null)
            {
                var singlePosLookups = parsedFont.GposTable.LookupList.Lookups
                    .Where(l => l.LookupType == 1)
                    .ToList();

                if (singlePosLookups.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"✅ Subset has {singlePosLookups.Count} SinglePos lookup(s)");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("ℹ️ No SinglePos lookups in subset (may not be present in Roboto)");
                }
            }
        }

        [TestMethod]
        public void Subset_WithGposMarkToBase_ShouldPreserveAccents()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Include base characters and their accented versions
            var chars = new[] {
        'e', 'a', 'o', 'u', 'i', 'n',
        'é', 'à', 'ö', 'ü', 'ñ', ' ',
        'E', 'A', 'O', 'U'
    };

            var subsetFont = font.CreateSubset(chars);

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_with_gpos_accents.ttf", parsedFont);

            // Assert
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

            // Check if MarkToBase lookups exist
            if (parsedFont.GposTable != null)
            {
                var markToBaseLookups = parsedFont.GposTable.LookupList.Lookups
                    .Where(l => l.LookupType == 4)
                    .ToList();

                if (markToBaseLookups.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"✅ Subset has {markToBaseLookups.Count} MarkToBase lookup(s)");

                    // Try to verify a specific attachment
                    var subtable = markToBaseLookups[0].SubTables.FirstOrDefault() as MarkToBaseSubTableFormat1;
                    if (subtable != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"   MarkClassCount: {subtable.MarkClassCount}");
                        System.Diagnostics.Debug.WriteLine($"   Marks: {subtable.MarkArray?.MarkCount ?? 0}");
                        System.Diagnostics.Debug.WriteLine($"   Bases: {subtable.BaseArray?.BaseCount ?? 0}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("ℹ️ No MarkToBase lookups in subset (may not be present in Roboto)");
                }
            }
        }

        [TestMethod]
        public void Subset_CompleteGposTest_AllThreeLookupTypes()
        {
            // Arrange - Kitchen sink test with all GPOS features
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false);

            // Characters covering:
            // - Kerning pairs (A-V, T-o, etc.)
            // - Potential single adjustments
            // - Accented characters for MarkToBase
            var text = "AVTO Wave Typography TEST café résumé ñoño 123";
            var subsetFont = font.CreateSubset(text);

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes, font.Format);

            // Save for inspection
            SaveFont("subset_Roboto_complete_gpos_test.ttf", parsedFont);

            // Assert
            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

            // Comprehensive GPOS check
            if (parsedFont.GposTable != null)
            {
                System.Diagnostics.Debug.WriteLine("=== GPOS Subsetting Results ===");
                System.Diagnostics.Debug.WriteLine($"Total lookups: {parsedFont.GposTable.LookupList.Lookups.Count}");

                var lookupsByType = parsedFont.GposTable.LookupList.Lookups
                    .GroupBy(l => l.LookupType)
                    .ToDictionary(g => g.Key, g => g.Count());

                foreach (var kvp in lookupsByType)
                {
                    string typeName = kvp.Key switch
                    {
                        1 => "SinglePos",
                        2 => "PairPos (Kerning)",
                        4 => "MarkToBase",
                        _ => $"Type {kvp.Key}"
                    };
                    System.Diagnostics.Debug.WriteLine($"  {typeName}: {kvp.Value} lookup(s)");
                }

                Assert.IsTrue(parsedFont.GposTable.LookupList.Lookups.Count > 0,
                    "Should have at least some GPOS lookups preserved");
            }
        }
    }
}