using EPPlus.Fonts.OpenType.Tables.Kern;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Provides kerning from legacy 'kern' table (pre-OpenType).
    /// Used as fallback when GPOS is not available.
    /// </summary>
    internal class LegacyKerningProvider
    {
        private readonly Dictionary<ulong, short> _kerningPairs;

        public LegacyKerningProvider(KernTable kernTable)
        {
            _kerningPairs = BuildKerningDictionary(kernTable);
        }

        /// <summary>
        /// Gets kerning adjustment for a glyph pair from kern table.
        /// O(1) lookup via pre-built dictionary.
        /// </summary>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            ulong key = MakeKey(leftGlyph, rightGlyph);

            if (_kerningPairs.TryGetValue(key, out short value))
                return value;

            return 0;
        }

        /// <summary>
        /// Builds a dictionary of all kerning pairs from the kern table.
        /// Called once during construction.
        /// </summary>
        private Dictionary<ulong, short> BuildKerningDictionary(KernTable kernTable)
        {
            var pairs = new Dictionary<ulong, short>();

            if (kernTable?.SubTables == null)
                return pairs;

            foreach (var subtable in kernTable.SubTables)
            {
                // Only support Format 0 (horizontal kerning)
                if (subtable.coverage.Format == 0 && subtable.Format0Subtable != null)
                {
                    AddPairsFromFormat0(pairs, subtable.Format0Subtable);
                }
            }

            return pairs;
        }

        /// <summary>
        /// Adds all kerning pairs from a Format 0 subtable to the dictionary.
        /// </summary>
        private void AddPairsFromFormat0(
            Dictionary<ulong, short> pairs,
            KernSubTableFormat0 format0)
        {
            if (format0.Pairs == null)
                return;

            foreach (var pair in format0.Pairs)
            {
                ulong key = MakeKey(pair.left, pair.right);

                // Last value wins if duplicate keys
                // (matches original behavior of returning first non-zero)
                if (pair.value != 0)
                {
                    pairs[key] = pair.value;
                }
            }
        }

        private static ulong MakeKey(ushort leftGlyph, ushort rightGlyph)
        {
            // Combine two ushorts into one ulong for fast dictionary lookup
            return ((ulong)leftGlyph << 16) | rightGlyph;
        }
    }
}