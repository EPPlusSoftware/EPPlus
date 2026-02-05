/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Type 2 (PairPos) handler
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Handlers
{
    /// <summary>
    /// Handler for GPOS Lookup Type 2: Pair Adjustment Positioning (Kerning).
    /// Subsets PairPos lookups by filtering glyph pairs and remapping IDs.
    /// </summary>
    internal class PairPosHandler : IGposLookupHandler
    {
        private class GlyphMapping
        {
            public ushort OldGlyphId { get; set; }
            public ushort NewGlyphId { get; set; }
            public int OldCoverageIndex { get; set; }
        }

        public ushort LookupType => 2;

        /// <summary>
        /// Phase 1: Discover.
        /// PairPos doesn't add new glyphs - it only positions existing pairs.
        /// </summary>
        public void Discover(FontSubsettingContext context, LookupTable lookup, GposSubsetProcessor processor)
        {
            // No-op: positioning doesn't require additional glyphs
        }

        /// <summary>
        /// Phase 2: Rewrite the lookup with subsetted data and remapped glyph IDs.
        /// </summary>
        public LookupTable Rewrite(FontSubsettingContext context, LookupTable lookup)
        {
            var newLookup = new LookupTable
            {
                LookupType = lookup.LookupType,
                LookupFlag = lookup.LookupFlag,
                MarkFilteringSet = lookup.MarkFilteringSet,
                SubTables = new List<FontTableElement>()
            };

            foreach (var subtable in lookup.SubTables)
            {
                GposSubTableBase rewritten = null;

                if (subtable is PairPosSubTableFormat1 format1)
                {
                    rewritten = RewriteFormat1(context, format1);
                }
                else if (subtable is PairPosSubTableFormat2 format2) // ← Aktivera!
                {
                    rewritten = RewriteFormat2(context, format2);
                }

                if (rewritten != null)
                {
                    newLookup.SubTables.Add(rewritten);
                }
            }

            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }

        /// <summary>
        /// Rewrites a PairPos Format 1 subtable.
        /// </summary>
        private PairPosSubTableFormat1 RewriteFormat1(FontSubsettingContext context, PairPosSubTableFormat1 original)
        {
            // Build list of included first glyphs with their coverage indices
            var includedGlyphs = new List<GlyphMapping>();

            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                int coverageIndex = original.Coverage.GetGlyphIndex(oldGlyphId);
                if (coverageIndex >= 0 && coverageIndex < original.PairSets.Count)
                {
                    if (context.OldToNewGlyphId.TryGetValue(oldGlyphId, out ushort newGlyphId))
                    {
                        includedGlyphs.Add(new GlyphMapping
                        {
                            OldGlyphId = oldGlyphId,
                            NewGlyphId = newGlyphId,
                            OldCoverageIndex = coverageIndex
                        });
                    }
                }
            }

            if (includedGlyphs.Count == 0)
                return null;

            // Sort by NEW glyph ID (for coverage)
            includedGlyphs.Sort((a, b) => a.NewGlyphId.CompareTo(b.NewGlyphId));

            // Build PairSets in new coverage order
            var newPairSets = new List<PairSet>();
            var newGlyphIds = new List<ushort>();

            foreach (var mapping in includedGlyphs)
            {
                var oldPairSet = original.PairSets[mapping.OldCoverageIndex];
                if (oldPairSet != null)
                {
                    var newPairSet = FilterPairSet(context, oldPairSet);

                    if (newPairSet != null && newPairSet.PairValueRecords.Count > 0)
                    {
                        newGlyphIds.Add(mapping.NewGlyphId);
                        newPairSets.Add(newPairSet);
                    }
                }
            }

            if (newPairSets.Count == 0)
                return null;

            // Create new coverage with sorted glyph IDs
            var newCoverage = new CoverageTableFormat1
            {
                CoverageFormat = 1,
                GlyphCount = (ushort)newGlyphIds.Count,
                GlyphArray = newGlyphIds.ToArray()
            };

            var rewritten = new PairPosSubTableFormat1
            {
                SubtableFormat = 1,
                ValueFormat1 = original.ValueFormat1,
                ValueFormat2 = original.ValueFormat2,
                Coverage = newCoverage,
                PairSets = newPairSets
            };

            return rewritten;
        }

        /// <summary>
        /// Rewrites a PairPos Format 2 subtable by expanding class-based kerning to Format 1.
        /// This simplifies subsetting and ensures compatibility.
        /// </summary>
        private PairPosSubTableFormat1 RewriteFormat2(FontSubsettingContext context, PairPosSubTableFormat2 original)
        {
            // Strategy: Expand class matrix to individual pairs (Format 1)
            // This is simpler than subsetting the class definitions and matrix

            // Build map of first glyphs that are both:
            // 1. In the original coverage
            // 2. In our subset
            var firstGlyphMappings = new List<GlyphMapping>();

            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                // Must be in coverage
                int coverageIndex = original.Coverage.GetGlyphIndex(oldGlyphId);
                if (coverageIndex < 0)
                    continue;

                // Must be in subset
                if (!context.OldToNewGlyphId.TryGetValue(oldGlyphId, out ushort newGlyphId))
                    continue;

                firstGlyphMappings.Add(new GlyphMapping
                {
                    OldGlyphId = oldGlyphId,
                    NewGlyphId = newGlyphId,
                    OldCoverageIndex = -1 // Not used for Format 2
                });
            }

            if (firstGlyphMappings.Count == 0)
                return null;

            // Sort by NEW glyph ID for coverage
            firstGlyphMappings.Sort((a, b) => a.NewGlyphId.CompareTo(b.NewGlyphId));

            // Build PairSets by expanding class matrix
            var newPairSets = new List<PairSet>();
            var newCoverageGlyphs = new List<ushort>();

            foreach (var firstMapping in firstGlyphMappings)
            {
                // Get class for first glyph
                int class1 = original.ClassDef1.GetClass(firstMapping.OldGlyphId);
                if (class1 < 0 || class1 >= original.Class1Count)
                    continue;

                // Collect all pairs for this first glyph
                var pairRecords = new List<PairValueRecord>();

                // Check all possible second glyphs in our subset
                foreach (var oldSecondGlyph in context.IncludedGlyphs)
                {
                    // Skip if same glyph
                    if (oldSecondGlyph == firstMapping.OldGlyphId)
                        continue;

                    // Get class for second glyph
                    int class2 = original.ClassDef2.GetClass(oldSecondGlyph);
                    if (class2 < 0 || class2 >= original.Class2Count)
                        continue;

                    // Look up in class matrix
                    var record = original.ClassMatrix[class1, class2];
                    if (record == null)
                        continue;

                    // Skip if no actual adjustment
                    bool hasValue1 = record.Value1 != null && record.Value1.XAdvance != 0;
                    bool hasValue2 = record.Value2 != null && record.Value2.XAdvance != 0;

                    if (!hasValue1 && !hasValue2)
                        continue;

                    // Get remapped second glyph ID
                    if (!context.OldToNewGlyphId.TryGetValue(oldSecondGlyph, out ushort newSecondGlyph))
                        continue;

                    // Add pair with remapped ID
                    pairRecords.Add(new PairValueRecord
                    {
                        SecondGlyph = newSecondGlyph,
                        Value1 = record.Value1,
                        Value2 = record.Value2
                    });
                }

                // Only include if we found pairs
                if (pairRecords.Count > 0)
                {
                    newCoverageGlyphs.Add(firstMapping.NewGlyphId);
                    newPairSets.Add(new PairSet
                    {
                        PairValueRecords = pairRecords
                    });
                }
            }

            if (newPairSets.Count == 0)
                return null;

            // Create new coverage
            var newCoverage = new CoverageTableFormat1
            {
                CoverageFormat = 1,
                GlyphCount = (ushort)newCoverageGlyphs.Count,
                GlyphArray = newCoverageGlyphs.ToArray()
            };

            // Return as Format 1 (expanded from class-based)
            return new PairPosSubTableFormat1
            {
                SubtableFormat = 1,
                ValueFormat1 = original.ValueFormat1,
                ValueFormat2 = original.ValueFormat2,
                Coverage = newCoverage,
                PairSets = newPairSets
            };
        }

        /// <summary>
        /// Filters a PairSet to only include pairs where second glyph is in subset.
        /// Remaps second glyph IDs.
        /// </summary>
        private PairSet FilterPairSet(FontSubsettingContext context, PairSet original)
        {
            var newRecords = new List<PairValueRecord>();

            foreach (var record in original.PairValueRecords)
            {
                // Check if second glyph is included in subset
                if (context.OldToNewGlyphId.TryGetValue(record.SecondGlyph, out ushort newSecondGlyph))
                {
                    var newRecord = new PairValueRecord
                    {
                        SecondGlyph = newSecondGlyph, // Remapped ID
                        Value1 = record.Value1,
                        Value2 = record.Value2
                    };
                    newRecords.Add(newRecord);
                }
            }

            if (newRecords.Count == 0)
                return null;

            return new PairSet
            {
                PairValueRecords = newRecords
            };
        }
    }
}