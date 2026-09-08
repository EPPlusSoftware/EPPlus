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
  09/07/2026         EPPlus Software AB           Fold script/language into the cache key
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Caches kerning values for glyph pairs to avoid repeated lookups.
    ///
    /// Keyed by (leftGlyph, rightGlyph, script, language) rather than just the glyph pair: the
    /// same pair can kern differently depending on which script's GPOS features are active, and a
    /// single cache instance is reused across many Shape() calls that may each specify a different
    /// script. Keying on the glyph pair alone would let a value computed for one script leak into
    /// a lookup for another.
    /// </summary>
    internal class KerningCache
    {
        private readonly Dictionary<string, short> _cache;

        public KerningCache()
        {
            _cache = new Dictionary<string, short>();
        }

        /// <summary>
        /// Tries to get a cached kerning value.
        /// </summary>
        public bool TryGet(ushort leftGlyph, ushort rightGlyph, string script, string language, out short value)
        {
            string key = MakeKey(leftGlyph, rightGlyph, script, language);
            return _cache.TryGetValue(key, out value);
        }

        /// <summary>
        /// Caches a kerning value.
        /// </summary>
        public void Set(ushort leftGlyph, ushort rightGlyph, string script, string language, short value)
        {
            string key = MakeKey(leftGlyph, rightGlyph, script, language);
            _cache[key] = value;
        }

        /// <summary>
        /// Clears the cache.
        /// </summary>
        public void Clear()
        {
            _cache.Clear();
        }

        private static string MakeKey(ushort leftGlyph, ushort rightGlyph, string script, string language)
        {
            // A string key trades a little speed for simplicity over the previous packed-ulong
            // key, which had no room left for script/language without a second dictionary layer.
            return leftGlyph.ToString() + "_" + rightGlyph.ToString() + "_"
                + (script ?? string.Empty) + "_" + (language ?? string.Empty);
        }
    }
}