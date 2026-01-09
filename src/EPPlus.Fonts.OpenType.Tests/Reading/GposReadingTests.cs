using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Reading
{
    [TestClass]
    public class GposReadingTests : FontTestBase
    {
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

        /*
         * Debug test, continue on Monday

        [TestMethod]
        public void Debug_InspectGposStructure()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);

            Assert.IsNotNull(pairPosLookup, "Need PairPos lookup");

            Trace.WriteLine($"=== GPOS Type 2 Lookup ===");
            Trace.WriteLine($"SubTables: {pairPosLookup.SubTables.Count}");

            // Inspect each subtable
            for (int i = 0; i < pairPosLookup.SubTables.Count; i++)
            {
                var subtable = pairPosLookup.SubTables[i] as PairPosSubTableFormat1;

                if (subtable == null)
                {
                    Trace.WriteLine($"SubTable[{i}]: Not Format1");
                    continue;
                }

                Trace.WriteLine($"\n=== SubTable[{i}] ===");
                Trace.WriteLine($"Format: {subtable.SubtableFormat}");
                Trace.WriteLine($"ValueFormat1: 0x{subtable.ValueFormat1:X4}");
                Trace.WriteLine($"ValueFormat2: 0x{subtable.ValueFormat2:X4}");
                Trace.WriteLine($"PairSets: {subtable.PairSets?.Count ?? 0}");

                // Check if Coverage has A and V
                if (font.CmapTable.TryGetGlyphId('A', out ushort aGlyph) &&
                    font.CmapTable.TryGetGlyphId('V', out ushort vGlyph))
                {
                    var coveredGlyphs = subtable.Coverage?.GetCoveredGlyphs() ?? new ushort[0];

                    bool aInCoverage = coveredGlyphs.Contains(aGlyph);
                    bool vInCoverage = coveredGlyphs.Contains(vGlyph);

                    Trace.WriteLine($"A({aGlyph}) in coverage: {aInCoverage}");
                    Trace.WriteLine($"V({vGlyph}) in coverage: {vInCoverage}");

                    // If A is in coverage, check its PairSet
                    if (aInCoverage)
                    {
                        int coverageIndex = subtable.Coverage.GetCoverageIndex(aGlyph);
                        Trace.WriteLine($"A coverage index: {coverageIndex}");

                        if (coverageIndex >= 0 && coverageIndex < subtable.PairSets.Count)
                        {
                            var pairSet = subtable.PairSets[coverageIndex];

                            if (pairSet != null && pairSet.PairValueRecords != null)
                            {
                                Trace.WriteLine($"A's PairSet has {pairSet.PairValueRecords.Count} pairs");

                                // List first 10 pairs
                                int count = 0;
                                foreach (var pair in pairSet.PairValueRecords.Take(10))
                                {
                                    Trace.WriteLine($"  Pair {count++}: SecondGlyph={pair.SecondGlyph}, XAdv={pair.Value1?.XAdvance ?? 0}");
                                }

                                // Check if V is there
                                bool hasV = pairSet.PairValueRecords.Any(p => p.SecondGlyph == vGlyph);
                                Trace.WriteLine($"V({vGlyph}) in A's PairSet: {hasV}");
                            }
                            else
                            {
                                Trace.WriteLine("A's PairSet is null or empty");
                            }
                        }
                    }
                }
            }
        }
        */

        [TestMethod]
        public void ReadGposTable_FindKerningPair_AV()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            // Get glyph IDs for 'A' and 'V'
            if (!font.CmapTable.TryGetGlyphId('A', out ushort aGlyph) ||
                !font.CmapTable.TryGetGlyphId('V', out ushort vGlyph))
            {
                Assert.Inconclusive("Could not get glyph IDs for A and V");
                return;
            }

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);
            Assert.IsNotNull(pairPosLookup, "Need PairPos lookup");

            var subtable = pairPosLookup.SubTables[0] as PairPosSubTableFormat1;
            Assert.IsNotNull(subtable, "Need PairPosSubTableFormat1");

            // Act - Try to get kerning for A+V
            bool found = subtable.TryGetPairAdjustment(aGlyph, vGlyph,
                out var value1, out var value2);

            // Assert
            Assert.IsTrue(found, $"Should find kerning pair for A({aGlyph}) + V({vGlyph})");
            Assert.IsNotNull(value1, "Value1 should exist");

            // Kerning for AV should be negative (tighter spacing)
            Assert.IsTrue(value1.XAdvance < 0,
                $"AV kerning should be negative (was {value1.XAdvance})");

            Trace.WriteLine($"A+V kerning: {value1.XAdvance} units");
        }

        [TestMethod]
        public void ReadGposTable_MultipleKerningPairs()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            var testPairs = new[]
            {
                ('A', 'V'),
                ('A', 'W'),
                ('A', 'Y'),
                ('T', 'o'),
                ('V', 'A')
            };

            var gpos = font.GposTable;
            var pairPosLookup = gpos.LookupList.Lookups.FirstOrDefault(l => l.LookupType == 2);
            var subtable = pairPosLookup?.SubTables[0] as PairPosSubTableFormat1;

            Assert.IsNotNull(subtable, "Need PairPosSubTableFormat1");

            int foundCount = 0;

            // Act - Check each pair
            foreach (var (first, second) in testPairs)
            {
                if (!font.CmapTable.TryGetGlyphId(first, out ushort gid1) ||
                    !font.CmapTable.TryGetGlyphId(second, out ushort gid2))
                    continue;

                if (subtable.TryGetPairAdjustment(gid1, gid2, out var val1, out var val2))
                {
                    foundCount++;
                    Trace.WriteLine($"{first}+{second}: XAdvance={val1.XAdvance}");
                }
            }

            // Assert
            Assert.IsTrue(foundCount >= 3,
                $"Should find at least 3 kerning pairs (found {foundCount})");
        }
    }
}