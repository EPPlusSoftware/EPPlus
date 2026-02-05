using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Reading
{
    [TestClass]
    public class GposReadingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }


        [TestMethod]
        public void ReadGposTable_Roboto()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            // Assert GPOS exists
            Assert.IsNotNull(font.GposTable, "Roboto should have GPOS table");

            var gpos = font.GposTable;

            // Verify basic structure
            Assert.AreEqual((ushort)1, gpos.MajorVersion, "GPOS major version should be 1");
            Assert.IsNotNull(gpos.ScriptList, "GPOS should have ScriptList");
            Assert.IsNotNull(gpos.FeatureList, "GPOS should have FeatureList");
            Assert.IsNotNull(gpos.LookupList, "GPOS should have LookupList");
        }

        [TestMethod]
        public void ReadGposTable_HasKernFeature()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            Assert.IsNotNull(gpos, "GPOS table should exist");

            // Act - Find 'kern' feature
            var kernFeature = gpos.FeatureList.FeatureRecords
                .FirstOrDefault(f => f.FeatureTag.Value == "kern");

            // Assert
            Assert.IsNotNull(kernFeature, "GPOS should have 'kern' feature");
            Assert.IsNotNull(kernFeature.FeatureTable, "kern feature should have FeatureTable");
            Assert.IsTrue(kernFeature.FeatureTable.LookupListIndices.Length > 0,
                "kern feature should reference at least one lookup");
        }

        [TestMethod]
        public void ReadGposTable_HasPairPosLookup()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            Assert.IsNotNull(gpos, "GPOS table should exist");

            // Act - Find Type 2 (PairPos) lookup
            var pairPosLookup = gpos.LookupList.Lookups
                .FirstOrDefault(l => l.LookupType == 2);

            // Assert
            Assert.IsNotNull(pairPosLookup, "GPOS should have Type 2 (PairPos) lookup");
            Assert.IsTrue(pairPosLookup.SubTables.Count > 0,
                "PairPos lookup should have subtables");
        }

        [TestMethod]
        public void ReadGposTable_PairPosFormat1()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);

            Assert.IsNotNull(pairPosLookup, "Need PairPos lookup for test");

            // Act - Get first PairPos subtable
            var subtable = pairPosLookup.SubTables[0] as PairPosSubTableFormat1;

            // Assert
            Assert.IsNotNull(subtable, "First subtable should be PairPosSubTableFormat1");
            Assert.AreEqual((ushort)1, subtable.SubtableFormat, "Format should be 1");
            Assert.IsNotNull(subtable.Coverage, "Should have Coverage table");
            Assert.IsNotNull(subtable.PairSets, "Should have PairSets");
            Assert.IsTrue(subtable.PairSets.Count > 0, "Should have at least one PairSet");
        }




        [TestMethod]
        public void ReadGposTable_FindKerningPair_ActualPairs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            if (!font.CmapTable.TryGetGlyphId('A', out ushort aGlyph))
            {
                Assert.Inconclusive("Could not get glyph ID for A");
                return;
            }

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);
            Assert.IsNotNull(pairPosLookup, "Need PairPos lookup");

            var subtable = pairPosLookup.SubTables[0] as PairPosSubTableFormat1;
            Assert.IsNotNull(subtable, "Need PairPosSubTableFormat1");

            // Act - Test with pairs we KNOW exist from debug output
            // A(37) + glyph 35 = kerning -61
            bool found = subtable.TryGetPairAdjustment(aGlyph, 35, out var value1, out var value2);

            // Assert
            Assert.IsTrue(found, $"Should find kerning pair for A({aGlyph}) + glyph 35");
            Assert.IsNotNull(value1, "Value1 should exist");
            Assert.AreEqual(-61, value1.XAdvance, "Should have XAdvance = -61");

            Debug.WriteLine($"✅ A + 35: XAdvance = {value1.XAdvance}");
        }

        [TestMethod]
        public void SerializeCmapTable()
        {
            Debug.WriteLine("=== SerializeCmapTable ===");
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("cmap");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);

            // Check IdDelta[3]
            var subtable = font.CmapTable.SubTables[0] as CmapSubtable4;
            Debug.WriteLine($"IdDelta[3] = {subtable.IdDelta[3]}");

            font.CmapTable.TryGetGlyphId('A', out ushort aGlyph);
            Debug.WriteLine($"A = {aGlyph}");

            var cmapBytes = font.CmapTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, cmapBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, cmapBytes);
        }

        [TestMethod]
        public void ReadGposTable_MultipleKerningPairs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);
            var subtable = pairPosLookup?.SubTables[0] as PairPosSubTableFormat1;

            Assert.IsNotNull(subtable, "Need PairPosSubTableFormat1");

            // Test known pairs from debug output
            // NOTE: All glyph IDs adjusted for current Roboto-Regular.ttf version
            // 'A' = glyph 37, 'V' = glyph 58
            var testPairs = new[]
            {
                (37, 35, -61),   // A + glyph 35
                (37, 88, -17),   // A + glyph 88
                (37, 91, -33),   // A + glyph 91
                (58, 13, 20),    // V + glyph 13 (adjusted from 14)
                (58, 65, 17),    // V + glyph 65 (adjusted from 66)
            };

            int foundCount = 0;

            // Act - Check each pair
            foreach (var (first, second, expectedKern) in testPairs)
            {
                if (subtable.TryGetPairAdjustment((ushort)first, (ushort)second, out var val1, out var val2))
                {
                    foundCount++;
                    Assert.AreEqual(expectedKern, val1.XAdvance,
                        $"Pair {first}+{second} should have kerning {expectedKern}");
                    Debug.WriteLine($"✅ {first}+{second}: XAdvance={val1.XAdvance}");
                }
                else
                {
                    Assert.Fail($"Should find pair {first}+{second}");
                }
            }

            // Assert
            Assert.AreEqual(5, foundCount, "Should find all 5 test pairs");
        }

        [TestMethod]
        public void ReadGposTable_HasSinglePosLookup()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            Assert.IsNotNull(gpos, "GPOS table should exist");

            // Act - Find Type 1 (SinglePos) lookup
            var singlePosLookup = gpos.LookupList.Lookups
                .FirstOrDefault(l => l.LookupType == 1);

            // Assert
            if (singlePosLookup != null)
            {
                Assert.IsTrue(singlePosLookup.SubTables.Count > 0,
                    "SinglePos lookup should have subtables");
                Debug.WriteLine($"✅ Found SinglePos lookup with {singlePosLookup.SubTables.Count} subtable(s)");
            }
            else
            {
                Assert.Inconclusive("Roboto does not have SinglePos (Type 1) lookups - this is OK");
            }
        }

        [TestMethod]
        public void ReadGposTable_SinglePosFormat1()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var singlePosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 1);

            if (singlePosLookup == null)
            {
                Assert.Inconclusive("Roboto does not have SinglePos lookups");
                return;
            }

            // Act - Get first SinglePos subtable
            var subtable = singlePosLookup.SubTables.FirstOrDefault() as SinglePosSubTableFormat1;

            if (subtable == null)
            {
                Assert.Inconclusive("First subtable is not Format 1");
                return;
            }

            // Assert
            Assert.AreEqual((ushort)1, subtable.SubtableFormat, "Format should be 1");
            Assert.IsNotNull(subtable.Coverage, "Should have Coverage table");
            Assert.IsNotNull(subtable.Value, "Should have ValueRecord");

            Debug.WriteLine($"✅ SinglePos Format 1:");
            Debug.WriteLine($"   ValueFormat: 0x{subtable.ValueFormat:X4}");
            Debug.WriteLine($"   XPlacement: {subtable.Value.XPlacement}");
            Debug.WriteLine($"   YPlacement: {subtable.Value.YPlacement}");
            Debug.WriteLine($"   XAdvance: {subtable.Value.XAdvance}");
            Debug.WriteLine($"   YAdvance: {subtable.Value.YAdvance}");
        }

        [TestMethod]
        public void ReadGposTable_SinglePosFormat2()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var singlePosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 1);

            if (singlePosLookup == null)
            {
                Assert.Inconclusive("Roboto does not have SinglePos lookups");
                return;
            }

            // Act - Find Format 2 subtable
            var subtable = singlePosLookup.SubTables
                .OfType<SinglePosSubTableFormat2>()
                .FirstOrDefault();

            if (subtable == null)
            {
                Assert.Inconclusive("No Format 2 subtables found");
                return;
            }

            // Assert
            Assert.AreEqual((ushort)2, subtable.SubtableFormat, "Format should be 2");
            Assert.IsNotNull(subtable.Coverage, "Should have Coverage table");
            Assert.IsNotNull(subtable.Values, "Should have ValueRecords array");
            Assert.AreEqual(subtable.ValueCount, subtable.Values.Length,
                "ValueCount should match array length");

            Debug.WriteLine($"✅ SinglePos Format 2:");
            Debug.WriteLine($"   ValueFormat: 0x{subtable.ValueFormat:X4}");
            Debug.WriteLine($"   ValueCount: {subtable.ValueCount}");
            Debug.WriteLine($"   First value - XPlacement: {subtable.Values[0].XPlacement}, YPlacement: {subtable.Values[0].YPlacement}");
        }

        [TestMethod]
        public void ReadGposTable_SinglePosTryGetAdjustment()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var singlePosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 1);

            if (singlePosLookup == null)
            {
                Assert.Inconclusive("Roboto does not have SinglePos lookups");
                return;
            }

            // Act - Test TryGetAdjustment on first subtable
            var subtableFormat1 = singlePosLookup.SubTables.OfType<SinglePosSubTableFormat1>().FirstOrDefault();
            var subtableFormat2 = singlePosLookup.SubTables.OfType<SinglePosSubTableFormat2>().FirstOrDefault();

            bool foundAdjustment = false;

            if (subtableFormat1 != null)
            {
                // Try a few common glyph IDs
                for (ushort glyphId = 1; glyphId < 100; glyphId++)
                {
                    if (subtableFormat1.TryGetAdjustment(glyphId, out var value))
                    {
                        Debug.WriteLine($"✅ Format 1: Glyph {glyphId} has adjustment: XAdv={value.XAdvance}, YAdv={value.YAdvance}");
                        foundAdjustment = true;
                        break;
                    }
                }
            }

            if (subtableFormat2 != null && !foundAdjustment)
            {
                // Try a few common glyph IDs
                for (ushort glyphId = 1; glyphId < 100; glyphId++)
                {
                    if (subtableFormat2.TryGetAdjustment(glyphId, out var value))
                    {
                        Debug.WriteLine($"✅ Format 2: Glyph {glyphId} has adjustment: XAdv={value.XAdvance}, YAdv={value.YAdvance}");
                        foundAdjustment = true;
                        break;
                    }
                }
            }

            if (!foundAdjustment)
            {
                Assert.Inconclusive("No adjustments found in coverage - this is OK if font doesn't use SinglePos heavily");
            }
        }

        [TestMethod]
        public void ReadGposTable_HasMarkToBaseLookup()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            Assert.IsNotNull(gpos, "GPOS table should exist");

            // Act - Find Type 4 (MarkToBase) lookup
            var markToBaseLookup = gpos.LookupList.Lookups
                .FirstOrDefault(l => l.LookupType == 4);

            // Assert
            if (markToBaseLookup != null)
            {
                Assert.IsTrue(markToBaseLookup.SubTables.Count > 0,
                    "MarkToBase lookup should have subtables");
                Debug.WriteLine($"✅ Found MarkToBase lookup with {markToBaseLookup.SubTables.Count} subtable(s)");
            }
            else
            {
                Assert.Inconclusive("Roboto does not have MarkToBase (Type 4) lookups - this is OK");
            }
        }

        [TestMethod]
        public void ReadGposTable_MarkToBaseFormat1()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var markToBaseLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 4);

            if (markToBaseLookup == null)
            {
                Assert.Inconclusive("Roboto does not have MarkToBase lookups");
                return;
            }

            // Act - Get first MarkToBase subtable
            var subtable = markToBaseLookup.SubTables.FirstOrDefault() as MarkToBaseSubTableFormat1;

            if (subtable == null)
            {
                Assert.Inconclusive("First subtable is not MarkToBaseSubTableFormat1");
                return;
            }

            // Assert
            Assert.AreEqual((ushort)1, subtable.SubtableFormat, "Format should be 1");
            Assert.IsNotNull(subtable.MarkCoverage, "Should have MarkCoverage table");
            Assert.IsNotNull(subtable.BaseCoverage, "Should have BaseCoverage table");
            Assert.IsNotNull(subtable.MarkArray, "Should have MarkArray");
            Assert.IsNotNull(subtable.BaseArray, "Should have BaseArray");
            Assert.IsTrue(subtable.MarkClassCount > 0, "Should have at least one mark class");

            Debug.WriteLine($"✅ MarkToBase Format 1:");
            Debug.WriteLine($"   MarkClassCount: {subtable.MarkClassCount}");
            Debug.WriteLine($"   MarkArray count: {subtable.MarkArray.MarkCount}");
            Debug.WriteLine($"   BaseArray count: {subtable.BaseArray.BaseCount}");
        }

        [TestMethod]
        public void ReadGposTable_MarkToBaseTryGetAttachment()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var markToBaseLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 4);

            if (markToBaseLookup == null)
            {
                Assert.Inconclusive("Roboto does not have MarkToBase lookups");
                return;
            }

            var subtable = markToBaseLookup.SubTables.FirstOrDefault() as MarkToBaseSubTableFormat1;

            if (subtable == null)
            {
                Assert.Inconclusive("No MarkToBase subtable found");
                return;
            }

            // Act - Try to find an attachment
            // We need to know which glyphs are marks and bases
            // Let's try some common combinations

            bool foundAttachment = false;

            // Try first few marks with first few bases
            for (ushort markGlyph = 1; markGlyph < 100 && !foundAttachment; markGlyph++)
            {
                if (subtable.MarkCoverage.IsCovered(markGlyph))
                {
                    for (ushort baseGlyph = 1; baseGlyph < 100; baseGlyph++)
                    {
                        if (subtable.BaseCoverage.IsCovered(baseGlyph))
                        {
                            if (subtable.TryGetAttachment(markGlyph, baseGlyph, out var markAnchor, out var baseAnchor))
                            {
                                Debug.WriteLine($"✅ Found attachment:");
                                Debug.WriteLine($"   Mark glyph: {markGlyph}");
                                Debug.WriteLine($"   Base glyph: {baseGlyph}");
                                Debug.WriteLine($"   Mark anchor: ({markAnchor.XCoordinate}, {markAnchor.YCoordinate})");
                                Debug.WriteLine($"   Base anchor: ({baseAnchor.XCoordinate}, {baseAnchor.YCoordinate})");
                                foundAttachment = true;
                                break;
                            }
                        }
                    }
                }
            }

            if (!foundAttachment)
            {
                Assert.Inconclusive("No attachments found - this is OK if font doesn't use MarkToBase heavily");
            }
        }

        [TestMethod]
        public void ReadGposTable_MarkToBaseWithAccentedCharacters()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var markToBaseLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 4);

            if (markToBaseLookup == null)
            {
                Assert.Inconclusive("Roboto does not have MarkToBase lookups");
                return;
            }

            var subtable = markToBaseLookup.SubTables.FirstOrDefault() as MarkToBaseSubTableFormat1;

            if (subtable == null)
            {
                Assert.Inconclusive("No MarkToBase subtable found");
                return;
            }

            // Act - Try common accented character combinations
            // e + combining acute (́) = é
            // Get glyph IDs for these characters

            if (font.CmapTable.TryGetGlyphId('e', out ushort eGlyph))
            {
                Debug.WriteLine($"'e' = glyph {eGlyph}");

                // Combining acute accent Unicode: U+0301
                if (font.CmapTable.TryGetGlyphId('\u0301', out ushort acuteGlyph))
                {
                    Debug.WriteLine($"Combining acute = glyph {acuteGlyph}");

                    if (subtable.TryGetAttachment(acuteGlyph, eGlyph, out var markAnchor, out var baseAnchor))
                    {
                        Debug.WriteLine($"✅ Found 'é' attachment:");
                        Debug.WriteLine($"   Mark anchor: ({markAnchor.XCoordinate}, {markAnchor.YCoordinate})");
                        Debug.WriteLine($"   Base anchor: ({baseAnchor.XCoordinate}, {baseAnchor.YCoordinate})");
                    }
                    else
                    {
                        Debug.WriteLine("No attachment found for 'e' + acute");
                    }
                }
            }
        }
    }
}