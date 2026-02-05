/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/15/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Provides kerning from GPOS PairPos lookups (Type 2).
    /// Supports both Format 1 (individual pairs) and Format 2 (class-based).
    /// </summary>
    internal class GposKerningProvider
    {
        private readonly List<PairPosSubTable> _subtables;

        public GposKerningProvider(GposTable gposTable)
        {
            _subtables = FindAllKerningSubtables(gposTable);
        }

        /// <summary>
        /// Gets kerning adjustment for a glyph pair.
        /// Tries each subtable until a match is found.
        /// </summary>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            foreach (var subtable in _subtables)
            {
                if (subtable.TryGetPairAdjustment(leftGlyph, rightGlyph,
                    out var value1, out var value2))
                {
                    // Kerning is typically in value1.XAdvance
                    if (value1 != null && value1.XAdvance != 0)
                        return value1.XAdvance;
                }
            }

            return 0;
        }

        /// <summary>
        /// Finds all PairPos subtables in the "kern" feature.
        /// </summary>
        private List<PairPosSubTable> FindAllKerningSubtables(GposTable gpos)
        {
            var subtables = new List<PairPosSubTable>();

            if (gpos == null)
                return subtables;

            // Find "kern" feature
            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == "kern")
                {
                    var feature = featureRecord.FeatureTable;

                    // Get lookups for this feature
                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        if (lookupIndex >= gpos.LookupList.Lookups.Count)
                            continue;

                        var lookup = gpos.LookupList.Lookups[lookupIndex];

                        // We want PairPos (Type 2)
                        if (lookup.LookupType == 2)
                        {
                            // Collect ALL PairPos subtables (Format 1 and/or Format 2)
                            foreach (var subtable in lookup.SubTables)
                            {
                                if (subtable is PairPosSubTable pairPos)
                                {
                                    subtables.Add(pairPos);
                                }
                            }
                        }
                    }
                }
            }

            return subtables;
        }
    }
}