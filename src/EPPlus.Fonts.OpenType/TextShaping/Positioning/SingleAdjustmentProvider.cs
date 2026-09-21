/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Filter by ScriptList/LangSys, not just tag
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
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
        /// <summary>
        /// A subtable together with the FeatureList index it was recorded under. A plain struct
        /// is used instead of ValueTuple to avoid depending on the System.ValueTuple package on
        /// net462, which is not confirmed to be referenced by this project.
        /// </summary>
        private readonly struct IndexedSubtable
        {
            public readonly int FeatureIndex;
            public readonly GposSubTableBase Subtable;

            public IndexedSubtable(int featureIndex, GposSubTableBase subtable)
            {
                FeatureIndex = featureIndex;
                Subtable = subtable;
            }
        }

        private readonly OpenTypeFont _font;
        private readonly GposTable _gpos;

        // Feature tag -> every (FeatureList index, subtable) pair recorded under that tag, across
        // ALL scripts. The FeatureList index is what lets TryGetAdjustment filter down to the
        // ones the requested script can actually reach - unlike the tag alone, which collapses
        // every script's data for the same tag into one bucket.
        private readonly Dictionary<string, List<IndexedSubtable>> _subtablesByFeature;

        // Active-index sets are per (script, language), not per font, so they are cached
        // separately rather than recomputed on every call - a TextShaper instance is reused
        // across many Shape() calls that may each specify a different script.
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public SingleAdjustmentProvider(OpenTypeFont font)
        {
            _font = font;
            _gpos = font?.GposTable;
            _subtablesByFeature = new Dictionary<string, List<IndexedSubtable>>();

            if (_gpos != null)
            {
                BuildFeatureMap(_gpos);
            }
        }

        /// <summary>
        /// Tries to get positioning adjustment for a single glyph using specified features,
        /// restricted to the features reachable from the given script and language.
        /// </summary>
        /// <param name="glyphId">The glyph ID to look up</param>
        /// <param name="features">List of feature tags to search (e.g., [\"kern\"])</param>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered lookup, which
        /// reproduces the previous behavior for callers that have no script to give.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        /// <param name="value">The ValueRecord if found</param>
        /// <returns>True if an adjustment was found</returns>
        public bool TryGetAdjustment(ushort glyphId, List<string> features, string script, string language, out ValueRecord value)
        {
            if (features == null || features.Count == 0)
            {
                // No features specified - don't apply any single adjustments
                value = null;
                return false;
            }

            HashSet<int> activeIndices = GetActiveIndices(script, language);

            foreach (var feature in features)
            {
                if (!_subtablesByFeature.TryGetValue(feature, out var entries))
                {
                    continue;
                }

                foreach (var entry in entries)
                {
                    // null activeIndices means "no ScriptList to filter by" - keep every entry,
                    // matching the previous behavior rather than discarding features we cannot resolve.
                    if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                    {
                        continue;
                    }

                    if (TryGetAdjustmentFromSubtable(entry.Subtable, glyphId, out value))
                    {
                        return true;
                    }
                }
            }

            value = null;
            return false;
        }

        private HashSet<int> GetActiveIndices(string script, string language)
        {
            string cacheKey = (script ?? string.Empty) + "|" + (language ?? string.Empty);

            if (!_activeIndexCache.TryGetValue(cacheKey, out var indices))
            {
                indices = ScriptFeatureResolver.GetActiveFeatureIndices(_gpos?.ScriptList, script, language);
                _activeIndexCache[cacheKey] = indices;
            }

            return indices;
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
        /// Builds a map of feature tags to their Single Adjustment subtables, keeping each
        /// subtable's original FeatureList index so TryGetAdjustment can filter by script later.
        /// Unlike the previous version, entries are APPENDED rather than overwritten per tag: two
        /// FeatureRecords can legitimately share a tag (one per script), and both must survive so
        /// the script filter has something to choose between at lookup time.
        /// </summary>
        private void BuildFeatureMap(GposTable gpos)
        {
            if (gpos?.FeatureList == null || gpos.LookupList == null)
                return;

            var featureRecords = gpos.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];
                string featureTag = featureRecord.FeatureTag.Value;
                var feature = featureRecord.FeatureTable;

                if (feature?.LookupListIndices == null)
                    continue;

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
                                if (!_subtablesByFeature.TryGetValue(featureTag, out var list))
                                {
                                    list = new List<IndexedSubtable>();
                                    _subtablesByFeature[featureTag] = list;
                                }

                                list.Add(new IndexedSubtable(featureIndex, (GposSubTableBase)subtable));
                            }
                        }
                    }
                }
            }
        }
    }
}