/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Type 1 (SinglePos) handler
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Handlers
{
    /// <summary>
    /// Handler for GPOS Lookup Type 1: Single Adjustment Positioning.
    /// Subsets SinglePos lookups by filtering glyphs and remapping IDs.
    /// </summary>
    internal class SinglePosHandler : IGposLookupHandler
    {
        public ushort LookupType => 1;

        /// <summary>
        /// Phase 1: Discover.
        /// SinglePos doesn't add new glyphs - it only positions existing ones.
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

                if (subtable is SinglePosSubTableFormat1 format1)
                {
                    rewritten = RewriteFormat1(context, format1);
                }
                else if (subtable is SinglePosSubTableFormat2 format2)
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
        /// Rewrites a SinglePos Format 1 subtable.
        /// </summary>
        private SinglePosSubTableFormat1 RewriteFormat1(FontSubsettingContext context, SinglePosSubTableFormat1 original)
        {
            // Filter coverage to only included glyphs
            var newCoverage = FilterCoverage(context, original.Coverage);

            if (newCoverage == null)
                return null; // No glyphs remain

            var rewritten = new SinglePosSubTableFormat1
            {
                SubtableFormat = 1,
                ValueFormat = original.ValueFormat,
                Value = original.Value, // Same adjustment for all glyphs
                Coverage = newCoverage
            };

            return rewritten;
        }

        /// <summary>
        /// Rewrites a SinglePos Format 2 subtable.
        /// </summary>
        private SinglePosSubTableFormat2 RewriteFormat2(FontSubsettingContext context, SinglePosSubTableFormat2 original)
        {
            // Get list of included glyphs in coverage order
            var includedGlyphs = new List<ushort>();
            var includedValues = new List<ValueRecord>();

            for (ushort oldGlyphId = 0; oldGlyphId < 65535; oldGlyphId++)
            {
                int coverageIndex = original.Coverage.GetGlyphIndex(oldGlyphId);
                if (coverageIndex >= 0 && coverageIndex < original.Values.Length)
                {
                    // Check if this glyph is included in subset
                    if (context.OldToNewGlyphId.ContainsKey(oldGlyphId))
                    {
                        includedGlyphs.Add(context.OldToNewGlyphId[oldGlyphId]); // New ID
                        includedValues.Add(original.Values[coverageIndex]);
                    }
                }
            }

            if (includedGlyphs.Count == 0)
                return null; // No glyphs remain

            // Create new coverage with remapped glyph IDs
            var newCoverage = CreateCoverageFromGlyphs(includedGlyphs);

            var rewritten = new SinglePosSubTableFormat2
            {
                SubtableFormat = 2,
                ValueFormat = original.ValueFormat,
                ValueCount = (ushort)includedValues.Count,
                Values = includedValues.ToArray(),
                Coverage = newCoverage
            };

            return rewritten;
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