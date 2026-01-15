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
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Kern;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Provides kerning information from either GPOS (modern) or kern table (legacy).
    /// Handles caching and fallback logic.
    /// </summary>
    internal class KerningProvider
    {
        private readonly OpenTypeFont _font;
        private readonly Dictionary<ulong, short> _cache;
        private readonly bool _hasGpos;
        private readonly bool _hasKern;
        private PairPosSubTableFormat1 _gposKerningSubtable;

        public KerningProvider(OpenTypeFont font)
        {
            _font = font;
            _cache = new Dictionary<ulong, short>();
            _hasGpos = font.GposTable != null;
            _hasKern = font.KernTable != null;

            // Pre-locate GPOS kerning subtable if available
            if (_hasGpos)
            {
                _gposKerningSubtable = FindGposKerningSubtable();
            }
        }

        /// <summary>
        /// Gets kerning value for a glyph pair.
        /// Returns 0 if no kerning is defined.
        /// </summary>
        /// <param name="leftGlyph">Left glyph ID</param>
        /// <param name="rightGlyph">Right glyph ID</param>
        /// <returns>Kerning adjustment in font units (negative = closer)</returns>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            // Check cache first
            ulong key = MakeCacheKey(leftGlyph, rightGlyph);

            short cachedValue;
            if (_cache.TryGetValue(key, out cachedValue))
            {
                return cachedValue;
            }

            // Lookup kerning value
            short kernValue = LookupKerning(leftGlyph, rightGlyph);

            // Cache result (even if 0, to avoid repeated lookups)
            _cache[key] = kernValue;

            return kernValue;
        }

        /// <summary>
        /// Clears the kerning cache.
        /// Call this if you need to free memory.
        /// </summary>
        public void ClearCache()
        {
            _cache.Clear();
        }

        #region Private Methods

        private short LookupKerning(ushort leftGlyph, ushort rightGlyph)
        {
            // Try GPOS first (modern, preferred)
            if (_hasGpos && _gposKerningSubtable != null)
            {
                short gposKern = GetGposKerning(leftGlyph, rightGlyph);
                if (gposKern != 0)
                {
                    return gposKern;
                }
            }

            // Fallback to kern table (legacy)
            if (_hasKern)
            {
                return GetLegacyKerning(leftGlyph, rightGlyph);
            }

            return 0;
        }

        private short GetGposKerning(ushort leftGlyph, ushort rightGlyph)
        {
            if (_gposKerningSubtable == null)
                return 0;

            ValueRecord value1, value2;
            if (_gposKerningSubtable.TryGetPairAdjustment(leftGlyph, rightGlyph, out value1, out value2))
            {
                // Kerning is typically in value1.XAdvance
                return value1.XAdvance;
            }

            return 0;
        }

        private short GetLegacyKerning(ushort leftGlyph, ushort rightGlyph)
        {
            var kernTable = _font.KernTable;
            if (kernTable == null || kernTable.SubTables == null || kernTable.SubTables.Count == 0)
                return 0;

            // Iterate through subtables
            foreach (var subtable in kernTable.SubTables)
            {
                // Only support Format 0 (horizontal kerning)
                if (subtable.coverage.Format == 0 && subtable.Format0Subtable != null)
                {
                    short kernValue = GetKerningFromFormat0(subtable.Format0Subtable, leftGlyph, rightGlyph);
                    if (kernValue != 0)
                    {
                        return kernValue;
                    }
                }
            }

            return 0;
        }

        private short GetKerningFromFormat0(KernSubTableFormat0 format0, ushort leftGlyph, ushort rightGlyph)
        {
            if (format0.Pairs == null)
                return 0;

            // Linear search through kerning pairs
            // Note: Could be optimized with binary search if pairs are sorted
            foreach (var pair in format0.Pairs)
            {
                if (pair.left == leftGlyph && pair.right == rightGlyph)
                {
                    return pair.value;
                }
            }

            return 0;
        }

        private PairPosSubTableFormat1 FindGposKerningSubtable()
        {
            var gpos = _font.GposTable;
            if (gpos == null) return null;

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
                            // Return first PairPos Format 1 subtable
                            foreach (var subtable in lookup.SubTables)
                            {
                                var pairPos = subtable as PairPosSubTableFormat1;
                                if (pairPos != null)
                                {
                                    return pairPos;
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }

        private static ulong MakeCacheKey(ushort leftGlyph, ushort rightGlyph)
        {
            // Combine two ushorts into one ulong for fast dictionary lookup
            return ((ulong)leftGlyph << 16) | rightGlyph;
        }

        #endregion
    }
}