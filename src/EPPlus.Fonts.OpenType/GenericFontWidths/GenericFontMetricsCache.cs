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
    /// Process wide cache of the serialized font metrics loaded from
    /// Resources/TextMetrics.zip, and the single place that decides which font key a request
    /// resolves to.
    ///
    /// GenericFontMetricsTextMeasurerBase keeps its own static dictionary populated from
    /// <see cref="GenericFontMetricsLoader"/>. Adding a second consumer would otherwise mean a
    /// second copy of all font entries in memory, so both should be pointed at this cache.
    /// </summary>
    internal static class GenericFontMetricsCache
    {
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static readonly object _syncRoot = new object();

        /// <summary>
        /// All loaded font metrics, keyed by family and subfamily.
        /// </summary>
        internal static Dictionary<uint, SerializedFontMetrics> Fonts
        {
            get
            {
                if (_fonts == null)
                {
                    lock (_syncRoot)
                    {
                        if (_fonts == null)
                        {
                            _fonts = GenericFontMetricsLoader.LoadFontMetrics();
                        }
                    }
                }
                return _fonts;
            }
        }

        /// <summary>
        /// Returns the metrics for a font key, or null when the font has no serialized metrics.
        /// </summary>
        internal static SerializedFontMetrics GetMetrics(uint fontKey)
        {
            SerializedFontMetrics metrics;
            if (Fonts.TryGetValue(fontKey, out metrics))
            {
                return metrics;
            }
            return null;
        }

        /// <summary>
        /// Returns true when serialized metrics exist for the font key.
        /// </summary>
        internal static bool IsValidFont(uint fontKey)
        {
            return Fonts.ContainsKey(fontKey);
        }

        /// <summary>
        /// Resolves the key that should actually be measured against, falling back to the
        /// Regular subfamily of the same family when the requested subfamily has no metrics.
        /// Returns uint.MaxValue when the family is unknown entirely.
        ///
        /// Not every family ships all four subfamilies. Windows has no Arial Black Bold, no
        /// Impact Italic, no Calibri Light Bold and no Tahoma Italic, among others - fifteen
        /// combinations in total. Those used to be generated anyway, by taking the Regular
        /// font's advance widths and writing them out under the requested subfamily, which is
        /// why a third of the shipped library carried Regular's metrics. Now those files are
        /// simply absent, and the substitution happens here where it is visible instead of
        /// being baked into the data.
        ///
        /// The fallback deliberately does not widen anything for Bold. The generator's old
        /// attempt at that had no effect on the reported widths, so falling straight through to
        /// Regular reproduces the previous behaviour exactly. Applying a real bold correction
        /// would be a separate, measured change.
        /// </summary>
        internal static uint ResolveFontKey(uint requestedKey)
        {
            if (requestedKey == uint.MaxValue) return uint.MaxValue;
            if (Fonts.ContainsKey(requestedKey)) return requestedKey;

            // The low 16 bits hold the subfamily; clearing them gives the Regular variant.
            var regularKey = requestedKey & 0xFFFF0000;
            if (regularKey != requestedKey && Fonts.ContainsKey(regularKey))
            {
                return regularKey;
            }

            return uint.MaxValue;
        }
    }
}