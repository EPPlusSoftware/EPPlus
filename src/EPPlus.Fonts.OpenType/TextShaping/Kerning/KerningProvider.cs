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
  09/07/2026         EPPlus Software AB           Pass script/language through to GposKerningProvider
 *************************************************************************************************/
using System;
using System.Runtime.CompilerServices;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    /// <summary>
    /// Provides kerning adjustments for glyph pairs.
    /// Delegates to GPOS (modern) or legacy kern table.
    /// </summary>
    internal class KerningProvider
    {
        private readonly GposKerningProvider _gposProvider;
        private readonly LegacyKerningProvider _legacyProvider;
        private readonly KerningCache _cache;


        public KerningProvider(OpenTypeFont font)
        {
            _cache = new KerningCache();

            if (font.GposTable != null)
            {
                _gposProvider = new GposKerningProvider(font.GposTable);
            }

            if (font.KernTable != null)
                _legacyProvider = new LegacyKerningProvider(font.KernTable);
        }

        /// <summary>
        /// Gets kerning value for a glyph pair, restricted to the given script/language.
        /// Returns 0 if no kerning is defined.
        /// </summary>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered GPOS lookup.
        /// The legacy kern table has no script concept and is unaffected either way.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph, string script, string language)
        {
            // Cache key must include script: the same glyph pair can legitimately kern
            // differently (or not at all) depending on which script's GPOS features are active,
            // and a single KerningProvider instance is reused across many Shape() calls that may
            // each request a different script.
            if (_cache.TryGet(leftGlyph, rightGlyph, script, language, out short cachedValue))
                return cachedValue;

            short kernValue = LookupKerning(leftGlyph, rightGlyph, script, language);

            _cache.Set(leftGlyph, rightGlyph, script, language, kernValue);

            return kernValue;
        }

        public void ClearCache() => _cache.Clear();

        private short LookupKerning(ushort leftGlyph, ushort rightGlyph, string script, string language)
        {
            // Try GPOS first (modern, preferred)
            if (_gposProvider != null)
            {
                short gposKern = _gposProvider.GetKerning(leftGlyph, rightGlyph, script, language);
                if (gposKern != 0)
                    return gposKern;
            }

            // Fallback to kern table (legacy)
            if (_legacyProvider != null)
            {
                return _legacyProvider.GetKerning(leftGlyph, rightGlyph);
            }

            return 0;
        }
    }
}