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

        [TestMethod]
        public void Debug_CheckIfSameFontInstance()
        {
            var font1 = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);
            var font2 = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

            bool sameInstance = ReferenceEquals(font1, font2);
            Debug.WriteLine($"Same instance: {sameInstance}");

            font1.CmapTable.TryGetGlyphId('A', out ushort a1);
            font2.CmapTable.TryGetGlyphId('A', out ushort a2);

            Debug.WriteLine($"Font1: A = {a1}");
            Debug.WriteLine($"Font2: A = {a2}");
        }

        [TestMethod]
        public void Debug_CmapInRealTestScenario()
        {
            // Run this 10 times like MSTest does
            for (int run = 0; run < 10; run++)
            {
                Debug.WriteLine($"\n=== RUN {run} ===");

                // DON'T clear cache (like your real tests)
                // OpenTypeFonts.ClearFontCache();

                var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

                font.CmapTable.TryGetGlyphId('A', out ushort aGlyph);
                Debug.WriteLine($"Run {run}: A = {aGlyph}");

                Assert.AreEqual((ushort)38, aGlyph, $"Run {run} failed");
            }
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
            // A(38) + glyph 36 = kerning -61
            bool found = subtable.TryGetPairAdjustment(aGlyph, 36, out var value1, out var value2);

            // Assert
            Assert.IsTrue(found, $"Should find kerning pair for A({aGlyph}) + glyph 36");
            Assert.IsNotNull(value1, "Value1 should exist");
            Assert.AreEqual(-61, value1.XAdvance, "Should have XAdvance = -61");

            Debug.WriteLine($"✅ A + 36: XAdvance = {value1.XAdvance}");
        }

        [TestMethod]
        public void Debug_ParallelGsubAndGpos()
        {
            var results = new System.Collections.Concurrent.ConcurrentBag<(string test, ushort glyphId, bool success)>();

            var tasks = new System.Threading.Tasks.Task[20];

            for (int i = 0; i < 20; i++)
            {
                int taskId = i;
                bool isGpos = (i % 2 == 0); // Alternerande GSUB/GPOS

                tasks[i] = System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        // ❌ WITHOUT ignoreCache (like your real tests)
                        var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular);

                        string testType;
                        if (isGpos)
                        {
                            // Access GPOS (like your failing test)
                            var gpos = font.GposTable;
                            testType = "GPOS";
                        }
                        else
                        {
                            // Access GSUB
                            var gsub = font.GsubTable;
                            testType = "GSUB";
                        }

                        // Now check Cmap
                        font.CmapTable.TryGetGlyphId('A', out ushort aGlyph);

                        Debug.WriteLine($"[Task {taskId} - {testType}] A = {aGlyph}");

                        results.Add((testType, aGlyph, aGlyph == 38));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Task {taskId}] EXCEPTION: {ex.Message}");
                        results.Add(("ERROR", 0, false));
                    }
                });
            }

            System.Threading.Tasks.Task.WaitAll(tasks);

            // Analyze
            Debug.WriteLine("\n=== RESULTS ===");
            var gposResults = results.Where(r => r.test == "GPOS").ToList();
            var gsubResults = results.Where(r => r.test == "GSUB").ToList();

            Debug.WriteLine($"GPOS tests: {gposResults.Count}, Success: {gposResults.Count(r => r.success)}");
            Debug.WriteLine($"GSUB tests: {gsubResults.Count}, Success: {gsubResults.Count(r => r.success)}");

            var wrongGlyphs = results.Where(r => !r.success).ToList();
            if (wrongGlyphs.Any())
            {
                Debug.WriteLine("\nFAILURES:");
                foreach (var fail in wrongGlyphs)
                {
                    Debug.WriteLine($"  {fail.test}: A = {fail.glyphId}");
                }
            }

            // Assert
            Assert.IsTrue(results.All(r => r.success), "Some tasks got wrong glyph ID");
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
            var testPairs = new[]
            {
                (38, 36, -61),   // A + glyph 36
                (38, 89, -17),   // A + glyph 89
                (38, 92, -33),   // A + glyph 92
                (59, 14, 20),    // V + glyph 14
                (59, 66, 17),    // V + glyph 66
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
    }
}