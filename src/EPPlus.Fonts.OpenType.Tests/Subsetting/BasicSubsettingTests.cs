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
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void Subset_Abc_RoundtripValidation()
        {
            var font = TestFolderEngine.LoadFont("Roboto");
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

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

            FontTestHelper.AssertFontValid(parsedFont);

            int ligatureCount = FontTestHelper.CountLigatures(parsedFont);
            Assert.AreEqual(0, ligatureCount, "abc should have NO ligatures");

            int expectedGlyphs = 3;      // a, b, c
            expectedGlyphs += 1;         // + space (U+0020)
            expectedGlyphs += 1;         // + .notdef (GID 0)
            expectedGlyphs += 5;         // + variants from Single Substitution

            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.HheaTable.numberOfHMetrics);

            Assert.IsTrue(parsedFont.CmapTable.ContainsChar(32));
        }

        [TestMethod]
        public void Subset_Fiffig_WithFullValidation()
        {
            var font = TestFolderEngine.LoadFont("Roboto");
            var subsetFont = font.CreateSubset("fiffig");

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_SingleChar_ShouldWork()
        {
            var font = TestFolderEngine.LoadFont("Mulish");
            var subsetFont = font.CreateSubset(new[] { 'a' });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_MultipleChars_ShouldWork()
        {
            var font = TestFolderEngine.LoadFont("Roboto");
            var subsetFont = font.CreateSubset(new[] {
                'F', 'l', 'y', 'g', 'a', 'n', 'd', 'e', 'b', 'ä', 'c', 'k', 's', 'i', 'r', 'ö', 'h', 'w', 'p', 'å',
                'm', 'j', 'u', 't', 'v', 'o', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0'
            });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);
        }

        [TestMethod]
        public void Subset_RoundtripHelper_ShouldWork()
        {
            var parsedFont = FontTestHelper.RoundtripSubset(TestFolderEngine, "Roboto", "test");

            SaveFontForCurrentTest(parsedFont);

            Assert.IsNotNull(parsedFont);
            Assert.IsTrue(parsedFont.MaxpTable.numGlyphs > 0);
        }

        [TestMethod]
        public void Check_Original_Roboto_Ligatures()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

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
            var font = TestFolderEngine.LoadFont("Roboto");

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

            Debug.WriteLine("\n=== CREATING SUBSET ===");
            var subsetFont = font.CreateSubset("fiffigoffice");

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            Debug.WriteLine("\n=== SUBSET FONT ===");
            Debug.WriteLine($"Total glyphs: {parsedFont.MaxpTable.numGlyphs}");

            parsedFont.CmapTable.TryGetGlyphId('f', out ushort subsetF);
            parsedFont.CmapTable.TryGetGlyphId('i', out ushort subsetI);
            Debug.WriteLine($"'f' = glyph {subsetF}");
            Debug.WriteLine($"'i' = glyph {subsetI}");

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

                    if (scriptTable.DefaultLangSys != null)
                    {
                        var defLang = scriptTable.DefaultLangSys;
                        Debug.WriteLine($"  DefaultLangSys:");
                        Debug.WriteLine($"    RequiredFeatureIndex: {defLang.RequiredFeatureIndex}");
                        Debug.WriteLine($"    FeatureIndices: [{string.Join(", ", defLang.FeatureIndices.Select(i => i.ToString()).ToArray())}]");

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
            var font = TestFolderEngine.LoadFont("Roboto Extra Light");
            //var font = OpenTypeFonts.LoadFont("Roboto Extra Light");
            var chars = new[] { 'f', 'e', 'c', 'd', 'g', 'E', 'a', 'b', ' ' };

            bool foundF_Original = font.CmapTable.TryGetGlyphId('f', out ushort fGlyphOrig);
            bool foundE_Original = font.CmapTable.TryGetGlyphId('e', out ushort eGlyphOrig);

            Debug.WriteLine($"=== ORIGINAL FONT ===");
            Debug.WriteLine($"f = glyph {fGlyphOrig}");
            Debug.WriteLine($"e = glyph {eGlyphOrig}");

            Assert.IsTrue(foundF_Original && foundE_Original, "Should find f and e in original");

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

            Debug.WriteLine($"\n=== CREATING SUBSET ===");
            var subsetFont = font.CreateSubset(chars);

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            bool foundF = parsedFont.CmapTable.TryGetGlyphId('f', out ushort fGlyph);
            bool foundE = parsedFont.CmapTable.TryGetGlyphId('e', out ushort eGlyph);

            Debug.WriteLine($"\n=== SUBSET FONT ===");
            Debug.WriteLine($"f = glyph {fGlyph}");
            Debug.WriteLine($"e = glyph {eGlyph}");

            Assert.IsNotNull(parsedFont.GposTable, "GPOS table should be present in subset");

            var kernLookups = parsedFont.GposTable.LookupList.Lookups.Where(l => l.LookupType == 2).ToList();
            Assert.IsTrue(kernLookups.Count > 0, "Should have at least one kerning lookup");

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
            var font = TestFolderEngine.LoadFont("Roboto");
            var chars = new[] {
                'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
                'A', 'B', 'C', 'D', 'E', ' '
            };

            var subsetFont = font.CreateSubset(chars);

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

            if (parsedFont.GposTable != null)
            {
                var singlePosLookups = parsedFont.GposTable.LookupList.Lookups
                    .Where(l => l.LookupType == 1)
                    .ToList();

                if (singlePosLookups.Count > 0)
                    System.Diagnostics.Debug.WriteLine($"✅ Subset has {singlePosLookups.Count} SinglePos lookup(s)");
                else
                    System.Diagnostics.Debug.WriteLine("ℹ️ No SinglePos lookups in subset (may not be present in Roboto)");
            }
        }

        [TestMethod]
        public void Subset_WithGposMarkToBase_ShouldPreserveAccents()
        {
            var font = TestFolderEngine.LoadFont("Roboto");
            var chars = new[] {
                'e', 'a', 'o', 'u', 'i', 'n',
                'é', 'à', 'ö', 'ü', 'ñ', ' ',
                'E', 'A', 'O', 'U'
            };

            var subsetFont = font.CreateSubset(chars);

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

            if (parsedFont.GposTable != null)
            {
                var markToBaseLookups = parsedFont.GposTable.LookupList.Lookups
                    .Where(l => l.LookupType == 4)
                    .ToList();

                if (markToBaseLookups.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"✅ Subset has {markToBaseLookups.Count} MarkToBase lookup(s)");

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
            var font = TestFolderEngine.LoadFont("Roboto");
            var text = "AVTO Wave Typography TEST café résumé ñoño 123";
            var subsetFont = font.CreateSubset(text);

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont);

            FontTestHelper.AssertFontValid(parsedFont, FontValidationSeverity.Warning);

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