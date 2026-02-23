using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Provides kerning from GPOS PairPos lookups (Type 2).
    /// Supports both Format 1 (individual pairs) and Format 2 (class-based).
    /// Uses lazy per-query lookup via TryGetPairAdjustment instead of
    /// pre-expanding all possible glyph pairs.
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
        /// Delegates to PairPosSubTable.TryGetPairAdjustment which does:
        ///   Format 1: coverage lookup + binary search in PairSet — O(log n)
        ///   Format 2: coverage lookup + class lookup + matrix index — O(1)
        /// Combined with KerningCache in KerningProvider, each unique pair
        /// is looked up at most once.
        /// </summary>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            for (int i = 0; i < _subtables.Count; i++)
            {
                if (_subtables[i].TryGetPairAdjustment(leftGlyph, rightGlyph,
                    out var value1, out var value2))
                {
                    if (value1 != null && value1.XAdvance != 0)
                        return value1.XAdvance;
                }
            }

            return 0;
        }

        private List<PairPosSubTable> FindAllKerningSubtables(GposTable gpos)
        {
            var subtables = new List<PairPosSubTable>();

            if (gpos == null)
                return subtables;

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == "kern")
                {
                    var feature = featureRecord.FeatureTable;

                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        if (lookupIndex >= gpos.LookupList.Lookups.Count)
                            continue;

                        var lookup = gpos.LookupList.Lookups[lookupIndex];

                        if (lookup.LookupType == 2)
                        {
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