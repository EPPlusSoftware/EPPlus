/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Type 4 (MarkToBase) handler
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Handlers
{
    /// <summary>
    /// Handler for GPOS Lookup Type 4: MarkToBase Attachment Positioning.
    /// Subsets MarkToBase lookups by filtering marks and bases, remapping IDs.
    /// </summary>
    internal class MarkToBaseHandler : IGposLookupHandler
    {
        public ushort LookupType => 4;

        /// <summary>
        /// Phase 1: Discover.
        /// MarkToBase doesn't add new glyphs - it only positions existing marks and bases.
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
                if (subtable is MarkToBaseSubTableFormat1 format1)
                {
                    var rewritten = RewriteFormat1(context, format1);
                    if (rewritten != null)
                    {
                        newLookup.SubTables.Add(rewritten);
                    }
                }
            }

            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }

        /// <summary>
        /// Rewrites a MarkToBase Format 1 subtable.
        /// </summary>
        private MarkToBaseSubTableFormat1 RewriteFormat1(FontSubsettingContext context, MarkToBaseSubTableFormat1 original)
        {
            // Filter mark coverage
            var newMarkCoverage = FilterCoverage(context, original.MarkCoverage);
            if (newMarkCoverage == null)
                return null; // No marks remain

            // Filter base coverage
            var newBaseCoverage = FilterCoverage(context, original.BaseCoverage);
            if (newBaseCoverage == null)
                return null; // No bases remain

            // Rebuild MarkArray
            var newMarkArray = FilterMarkArray(context, original);
            if (newMarkArray == null || newMarkArray.MarkCount == 0)
                return null;

            // Rebuild BaseArray
            var newBaseArray = FilterBaseArray(context, original);
            if (newBaseArray == null || newBaseArray.BaseCount == 0)
                return null;

            var rewritten = new MarkToBaseSubTableFormat1
            {
                SubtableFormat = 1,
                MarkClassCount = original.MarkClassCount,
                MarkCoverage = newMarkCoverage,
                BaseCoverage = newBaseCoverage,
                MarkArray = newMarkArray,
                BaseArray = newBaseArray
            };

            return rewritten;
        }

        /// <summary>
        /// Filters MarkArray to only include marks in the subset.
        /// </summary>
        private MarkArray FilterMarkArray(FontSubsettingContext context, MarkToBaseSubTableFormat1 original)
        {
            var includedRecords = new List<MarkRecord>();

            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                int coverageIndex = original.MarkCoverage.GetGlyphIndex(oldGlyphId);
                if (coverageIndex >= 0 && coverageIndex < original.MarkArray.MarkCount)
                {
                    // Check if this mark is included in subset
                    if (context.OldToNewGlyphId.ContainsKey(oldGlyphId))
                    {
                        includedRecords.Add(original.MarkArray.Records[coverageIndex]);
                    }
                }
            }

            if (includedRecords.Count == 0)
                return null;

            return new MarkArray
            {
                MarkCount = (ushort)includedRecords.Count,
                Records = includedRecords.ToArray()
            };
        }

        /// <summary>
        /// Filters BaseArray to only include bases in the subset.
        /// </summary>
        private BaseArray FilterBaseArray(FontSubsettingContext context, MarkToBaseSubTableFormat1 original)
        {
            var includedRecords = new List<BaseRecord>();

            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                int coverageIndex = original.BaseCoverage.GetGlyphIndex(oldGlyphId);
                if (coverageIndex >= 0 && coverageIndex < original.BaseArray.BaseCount)
                {
                    // Check if this base is included in subset
                    if (context.OldToNewGlyphId.ContainsKey(oldGlyphId))
                    {
                        includedRecords.Add(original.BaseArray.Records[coverageIndex]);
                    }
                }
            }

            if (includedRecords.Count == 0)
                return null;

            return new BaseArray
            {
                BaseCount = (ushort)includedRecords.Count,
                Records = includedRecords.ToArray()
            };
        }

        /// <summary>
        /// Filters a coverage table to only included glyphs and remaps IDs.
        /// </summary>
        private CoverageTable FilterCoverage(FontSubsettingContext context, CoverageTable original)
        {
            if (original == null)
                return null;

            var includedGlyphs = new List<ushort>();

            // Iterate through all glyphs in original coverage
            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                if (original.GetGlyphIndex(oldGlyphId) >= 0)
                {
                    // This glyph is in coverage - check if included in subset
                    if (context.OldToNewGlyphId.TryGetValue(oldGlyphId, out ushort newGlyphId))
                    {
                        includedGlyphs.Add(newGlyphId);
                    }
                }
            }

            if (includedGlyphs.Count == 0)
                return null;

            return CreateCoverageFromGlyphs(includedGlyphs);
        }

        /// <summary>
        /// Creates a Coverage Format 1 table from a list of glyph IDs.
        /// </summary>
        private CoverageTable CreateCoverageFromGlyphs(List<ushort> glyphIds)
        {
            // Sort glyphs for Coverage Format 1
            glyphIds.Sort();

            return new CoverageTableFormat1
            {
                CoverageFormat = 1,
                GlyphCount = (ushort)glyphIds.Count,
                GlyphArray = glyphIds.ToArray()
            };
        }
    }
}