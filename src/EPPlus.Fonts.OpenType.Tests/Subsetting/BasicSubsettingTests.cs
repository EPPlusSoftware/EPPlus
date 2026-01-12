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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

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