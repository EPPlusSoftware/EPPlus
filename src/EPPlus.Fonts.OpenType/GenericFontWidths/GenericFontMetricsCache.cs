/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/01/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.GenericFontWidths
{
    /// <summary>
    /// Process wide cache of the serialized font metrics, and the single place that decides
    /// which font key a request resolves to.
    ///
    /// Fonts are loaded on first use rather than all at once. The set of available keys is read
    /// from the archive's file names, so IsValidFont and ResolveFontKey can answer without
    /// decompressing anything; only fonts that are actually measured get parsed. A workbook
    /// using two fonts holds a few kB here rather than the whole library.
    /// </summary>
    internal static class GenericFontMetricsCache
    {
        private static readonly Dictionary<uint, SerializedFontMetrics> _loaded =
            new Dictionary<uint, SerializedFontMetrics>();
        private static HashSet<uint> _availableKeys;
        private static readonly object _syncRoot = new object();

        private static HashSet<uint> AvailableKeys
        {
            get
            {
                if (_availableKeys == null)
                {
                    lock (_syncRoot)
                    {
                        if (_availableKeys == null)
                        {
                            _availableKeys = GenericFontMetricsLoader.LoadAvailableFontKeys();
                        }
                    }
                }
                return _availableKeys;
            }
        }

        /// <summary>
        /// Returns the metrics for a font key, or null when the archive holds no such font.
        /// </summary>
        internal static SerializedFontMetrics GetMetrics(uint fontKey)
        {
            SerializedFontMetrics metrics;
            lock (_syncRoot)
            {
                if (_loaded.TryGetValue(fontKey, out metrics))
                {
                    return metrics;
                }
            }

            if (!AvailableKeys.Contains(fontKey)) return null;

            // Parsed outside the lock; a duplicate parse under contention is cheaper than
            // holding the lock across the decompression, and the result is identical either way.
            var parsed = GenericFontMetricsLoader.LoadFontMetrics(fontKey);
            if (parsed == null) return null;

            lock (_syncRoot)
            {
                if (_loaded.TryGetValue(fontKey, out metrics))
                {
                    return metrics;
                }
                _loaded[fontKey] = parsed;
                return parsed;
            }
        }

        /// <summary>
        /// True when the archive contains metrics for the font key. Does not load them.
        /// </summary>
        internal static bool IsValidFont(uint fontKey)
        {
            return AvailableKeys.Contains(fontKey);
        }

        /// <summary>
        /// Resolves the key that should actually be measured against, falling back to the
        /// Regular subfamily of the same family when the requested subfamily has no metrics.
        /// Returns uint.MaxValue when the family is unknown entirely.
        ///
        /// Not every family ships all four subfamilies. Windows has no Arial Black Bold, no
        /// Impact Italic, no Calibri Light Bold and no Tahoma Italic, among others - fifteen
        /// combinations in total. Those used to be generated anyway, by taking the Regular
        /// font's advance widths and writing them out under the requested subfamily. Now those
        /// files are simply absent, and the substitution happens here where it is visible
        /// instead of being baked into the data.
        ///
        /// The fallback deliberately does not widen anything for Bold. The generator's old
        /// attempt at that had no effect on the reported widths, so falling straight through to
        /// Regular reproduces the previous behaviour exactly.
        /// </summary>
        internal static uint ResolveFontKey(uint requestedKey)
        {
            if (requestedKey == uint.MaxValue) return uint.MaxValue;
            if (AvailableKeys.Contains(requestedKey)) return requestedKey;

            // The low 16 bits hold the subfamily; clearing them gives the Regular variant.
            var regularKey = requestedKey & 0xFFFF0000;
            if (regularKey != requestedKey && AvailableKeys.Contains(regularKey))
            {
                return regularKey;
            }

            return uint.MaxValue;
        }
    }
}