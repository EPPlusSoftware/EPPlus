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
        /// Gets kerning value for a glyph pair.
        /// Returns 0 if no kerning is defined.
        /// </summary>
        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            // Check cache first
            if (_cache.TryGet(leftGlyph, rightGlyph, out short cachedValue))
                return cachedValue;

            // Lookup kerning value
            short kernValue = LookupKerning(leftGlyph, rightGlyph);

            // Cache result
            _cache.Set(leftGlyph, rightGlyph, kernValue);

            return kernValue;
        }

        public void ClearCache() => _cache.Clear();

        private short LookupKerning(ushort leftGlyph, ushort rightGlyph)
        {
            // Try GPOS first (modern, preferred)
            if (_gposProvider != null)
            {
                short gposKern = _gposProvider.GetKerning(leftGlyph, rightGlyph);
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