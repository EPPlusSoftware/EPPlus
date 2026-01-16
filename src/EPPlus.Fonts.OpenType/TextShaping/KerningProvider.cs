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
    /// Supports both PairPos Format 1 and Format 2.
    /// </summary>
    internal class KerningProvider
    {
        private readonly OpenTypeFont _font;
        private readonly Dictionary<ulong, short> _cache;
        private readonly bool _hasGpos;
        private readonly bool _hasKern;
        private List<PairPosSubTable> _gposKerningSubtables;

        public KerningProvider(OpenTypeFont font)
        {
            _font = font;
            _cache = new Dictionary<ulong, short>();
            _hasGpos = font.GposTable != null;
            _hasKern = font.KernTable != null;

            // Pre-locate ALL GPOS kerning subtables
            if (_hasGpos)
            {
                _gposKerningSubtables = FindAllGposKerningSubtables();
            }
        }

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

        public void ClearCache()
        {
            _cache.Clear();
        }

        #region Private Methods

        private short LookupKerning(ushort leftGlyph, ushort rightGlyph)
        {
            // Try GPOS first (modern, preferred)
            if (_hasGpos && _gposKerningSubtables != null && _gposKerningSubtables.Count > 0)
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
            if (_gposKerningSubtables == null || _gposKerningSubtables.Count == 0)
                return 0;

            // Try each subtable until we find a match
            // This handles fonts with multiple subtables (Format1 + Format2)
            foreach (var subtable in _gposKerningSubtables)
            {
                ValueRecord value1, value2;
                if (subtable.TryGetPairAdjustment(leftGlyph, rightGlyph,
                    out value1, out value2))
                {
                    // Kerning is typically in value1.XAdvance
                    if (value1 != null && value1.XAdvance != 0)
                        return value1.XAdvance;
                }
            }

            return 0;
        }

        private short GetLegacyKerning(ushort leftGlyph, ushort rightGlyph)
        {
            var kernTable = _font.KernTable;
            if (kernTable == null || kernTable.SubTables == null || kernTable.SubTables.Count == 0)
                return 0;

            foreach (var subtable in kernTable.SubTables)
            {
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

            foreach (var pair in format0.Pairs)
            {
                if (pair.left == leftGlyph && pair.right == rightGlyph)
                {
                    return pair.value;
                }
            }

            return 0;
        }

        private List<PairPosSubTable> FindAllGposKerningSubtables()
        {
            var subtables = new List<PairPosSubTable>();
            var gpos = _font.GposTable;

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

        private static ulong MakeCacheKey(ushort leftGlyph, ushort rightGlyph)
        {
            return ((ulong)leftGlyph << 16) | rightGlyph;
        }

        #endregion
    }
}