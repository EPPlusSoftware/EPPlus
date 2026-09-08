/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Extension-wrapped kerning support
  09/07/2026         EPPlus Software AB           Filter by ScriptList/LangSys, not just tag
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Provides kerning from GPOS PairPos lookups (Type 2), including lookups that are
    /// extension wrapped (Type 9) in the font file and unwrapped by GposTableLoader.
    /// Supports both Format 1 (individual pairs) and Format 2 (class-based).
    /// Uses lazy per-query lookup via TryGetPairAdjustment instead of
    /// pre-expanding all possible glyph pairs.
    /// </summary>
    internal class GposKerningProvider
    {
        private readonly struct IndexedSubtable
        {
            public readonly int FeatureIndex;
            public readonly PairPosSubTable Subtable;

            public IndexedSubtable(int featureIndex, PairPosSubTable subtable)
            {
                FeatureIndex = featureIndex;
                Subtable = subtable;
            }
        }

        private readonly GposTable _gpos;
        private readonly List<IndexedSubtable> _subtables;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public GposKerningProvider(GposTable gposTable)
        {
            _gpos = gposTable;
            _subtables = FindAllKerningSubtables(gposTable);
        }

        /// <summary>
        /// Gets kerning adjustment for a glyph pair, restricted to the "kern" feature records
        /// reachable from the given script and language.
        /// </summary>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered lookup, which
        /// reproduces the previous behavior for callers that have no script to give.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph, string script, string language)
        {
            HashSet<int> activeIndices = GetActiveIndices(script, language);

            for (int i = 0; i < _subtables.Count; i++)
            {
                // null activeIndices means "no ScriptList to filter by" - keep every entry,
                // matching the previous behavior rather than discarding kerning we cannot resolve.
                if (activeIndices != null && !activeIndices.Contains(_subtables[i].FeatureIndex))
                {
                    continue;
                }

                if (_subtables[i].Subtable.TryGetPairAdjustment(leftGlyph, rightGlyph,
                    out var value1, out var value2))
                {
                    if (value1 != null && value1.XAdvance != 0)
                        return value1.XAdvance;
                }
            }

            return 0;
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

        /// <summary>
        /// Collects every "kern"-tagged PairPos subtable together with its original FeatureList
        /// index, across all scripts, so GetKerning can filter by script later. Two FeatureRecords
        /// can legitimately share the "kern" tag (one per script); both are kept.
        /// </summary>
        private List<IndexedSubtable> FindAllKerningSubtables(GposTable gpos)
        {
            var subtables = new List<IndexedSubtable>();

            if (gpos?.FeatureList == null)
                return subtables;

            var featureRecords = gpos.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];

                if (featureRecord.FeatureTag.Value != "kern")
                    continue;

                var feature = featureRecord.FeatureTable;

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex >= gpos.LookupList.Lookups.Count)
                        continue;

                    var lookup = gpos.LookupList.Lookups[lookupIndex];

                    // No lookup type check here. Extension positioning (type 9) is resolved
                    // by GposTableLoader, which stores the wrapped subtable directly in the
                    // lookup while leaving LookupType at 9. Requiring LookupType == 2 would
                    // therefore discard all extension wrapped kerning. The subtable type
                    // test below is what actually identifies pair positioning data.
                    foreach (var subtable in lookup.SubTables)
                    {
                        if (subtable is PairPosSubTable pairPos)
                        {
                            subtables.Add(new IndexedSubtable(featureIndex, pairPos));
                        }
                    }
                }
            }

            return subtables;
        }
    }
}