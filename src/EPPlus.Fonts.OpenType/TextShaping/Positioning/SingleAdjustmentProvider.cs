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
        private readonly OpenTypeFont _font;
        private readonly Dictionary<string, List<GposSubTableBase>> _subtablesByFeature;

        public SingleAdjustmentProvider(OpenTypeFont font)
        {
            _font = font;
            _subtablesByFeature = new Dictionary<string, List<GposSubTableBase>>();

            if (font?.GposTable != null)
            {
                BuildFeatureMap(font.GposTable);
            }
        }

        /// <summary>
        /// Tries to get positioning adjustment for a single glyph using specified features.
        /// </summary>
        /// <param name="glyphId">The glyph ID to look up</param>
        /// <param name="features">List of feature tags to search (e.g., ["kern"])</param>
        /// <param name="value">The ValueRecord if found</param>
        /// <returns>True if an adjustment was found</returns>
        public bool TryGetAdjustment(ushort glyphId, List<string> features, out ValueRecord value)
        {
            if (features == null || features.Count == 0)
            {
                // No features specified - don't apply any single adjustments
                value = null;
                return false;
            }

            // Search in specified features
            foreach (var feature in features)
            {
                if (_subtablesByFeature.TryGetValue(feature, out var subtables))
                {
                    foreach (var subtable in subtables)
                    {
                        if (TryGetAdjustmentFromSubtable(subtable, glyphId, out value))
                        {
                            return true;
                        }
                    }
                }
            }

            value = null;
            return false;
        }

        private bool TryGetAdjustmentFromSubtable(GposSubTableBase subtable, ushort glyphId, out ValueRecord value)
        {
            // Try Format 1
            if (subtable is SinglePosSubTableFormat1 format1)
            {
                return format1.TryGetAdjustment(glyphId, out value);
            }
            // Try Format 2
            else if (subtable is SinglePosSubTableFormat2 format2)
            {
                return format2.TryGetAdjustment(glyphId, out value);
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Builds a map of feature tags to their Single Adjustment subtables.
        /// </summary>
        private void BuildFeatureMap(GposTable gpos)
        {
            if (gpos?.FeatureList == null || gpos.LookupList == null)
                return;

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                string featureTag = featureRecord.FeatureTag.Value;
                var feature = featureRecord.FeatureTable;

                if (feature?.LookupListIndices == null)
                    continue;

                var subtables = new List<GposSubTableBase>();

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex >= gpos.LookupList.Lookups.Count)
                        continue;

                    var lookup = gpos.LookupList.Lookups[lookupIndex];

                    // We want Single Adjustment (Type 1)
                    if (lookup.LookupType == 1 && lookup.SubTables != null)
                    {
                        foreach (var subtable in lookup.SubTables)
                        {
                            if (subtable is SinglePosSubTableFormat1 || subtable is SinglePosSubTableFormat2)
                            {
                                subtables.Add((GposSubTableBase)subtable);
                            }
                        }
                    }
                }

                if (subtables.Count > 0)
                {
                    _subtablesByFeature[featureTag] = subtables;
                }
            }
        }
    }
}