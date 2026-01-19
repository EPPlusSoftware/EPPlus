/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/19/2026         EPPlus Software AB           GPOS Single Adjustment support
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Positioning
{
    /// <summary>
    /// Provides single glyph positioning adjustments from GPOS Lookup Type 1.
    /// Handles both Format 1 (uniform adjustment) and Format 2 (per-glyph adjustments).
    /// </summary>
    internal class SingleAdjustmentProvider
    {
        private readonly List<GposSubTableBase> _subtables;

        public SingleAdjustmentProvider(OpenTypeFont font)
        {
            _subtables = new List<GposSubTableBase>();

            if (font?.GposTable != null)
            {
                _subtables = FindAllSingleAdjustmentSubtables(font.GposTable);
            }
        }

        /// <summary>
        /// Tries to get positioning adjustment for a single glyph.
        /// </summary>
        /// <param name="glyphId">The glyph ID to look up</param>
        /// <param name="value">The ValueRecord if found</param>
        /// <returns>True if an adjustment was found</returns>
        public bool TryGetAdjustment(ushort glyphId, out ValueRecord value)
        {
            foreach (var subtable in _subtables)
            {
                // Try Format 1
                if (subtable is SinglePosSubTableFormat1 format1)
                {
                    if (format1.TryGetAdjustment(glyphId, out value))
                    {
                        return true;
                    }
                }
                // Try Format 2
                else if (subtable is SinglePosSubTableFormat2 format2)
                {
                    if (format2.TryGetAdjustment(glyphId, out value))
                    {
                        return true;
                    }
                }
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Finds all Single Adjustment subtables (Type 1) in the GPOS table.
        /// Collects from all relevant features (not just one specific feature tag).
        /// </summary>
        private List<GposSubTableBase> FindAllSingleAdjustmentSubtables(GposTable gpos)
        {
            var subtables = new List<GposSubTableBase>();

            if (gpos?.FeatureList == null || gpos.LookupList == null)
                return subtables;

            // Collect all lookup indices from all features
            var lookupIndices = new HashSet<ushort>();

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                var feature = featureRecord.FeatureTable;
                if (feature?.LookupListIndices != null)
                {
                    foreach (var index in feature.LookupListIndices)
                    {
                        lookupIndices.Add(index);
                    }
                }
            }

            // Process each unique lookup
            foreach (var lookupIndex in lookupIndices)
            {
                if (lookupIndex >= gpos.LookupList.Lookups.Count)
                    continue;

                var lookup = gpos.LookupList.Lookups[lookupIndex];

                // We want Single Adjustment (Type 1)
                if (lookup.LookupType == 1)
                {
                    // Collect ALL Single Adjustment subtables (Format 1 and/or Format 2)
                    if (lookup.SubTables != null)
                    {
                        foreach (var subtable in lookup.SubTables)
                        {
                            if (subtable is SinglePosSubTableFormat1 format1)
                            {
                                subtables.Add(format1);
                            }
                            else if (subtable is SinglePosSubTableFormat2 format2)
                            {
                                subtables.Add(format2);
                            }
                        }
                    }
                }
            }

            return subtables;
        }
    }
}