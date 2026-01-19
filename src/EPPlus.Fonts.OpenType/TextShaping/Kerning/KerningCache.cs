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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Caches kerning values for glyph pairs to avoid repeated lookups.
    /// </summary>
    internal class KerningCache
    {
        private readonly Dictionary<ulong, short> _cache;

        public KerningCache()
        {
            _cache = new Dictionary<ulong, short>();
        }

        /// <summary>
        /// Tries to get a cached kerning value.
        /// </summary>
        public bool TryGet(ushort leftGlyph, ushort rightGlyph, out short value)
        {
            ulong key = MakeKey(leftGlyph, rightGlyph);
            return _cache.TryGetValue(key, out value);
        }

        /// <summary>
        /// Caches a kerning value.
        /// </summary>
        public void Set(ushort leftGlyph, ushort rightGlyph, short value)
        {
            ulong key = MakeKey(leftGlyph, rightGlyph);
            _cache[key] = value;
        }

        /// <summary>
        /// Clears the cache.
        /// </summary>
        public void Clear()
        {
            _cache.Clear();
        }

        private static ulong MakeKey(ushort leftGlyph, ushort rightGlyph)
        {
            // Combine two ushorts into one ulong for fast dictionary lookup
            return ((ulong)leftGlyph << 16) | rightGlyph;
        }
    }
}