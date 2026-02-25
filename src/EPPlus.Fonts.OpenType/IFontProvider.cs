/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Provides fonts for text shaping with fallback support.
    /// </summary>
    public interface IFontProvider
    {
        /// <summary>
        /// Gets the primary font (user's chosen font).
        /// </summary>
        OpenTypeFont PrimaryFont { get; }

        /// <summary>
        /// Tries to find a font that contains the specified code point.
        /// Searches primary font first, then fallbacks.
        /// </summary>
        /// <param name="codePoint">Unicode code point</param>
        /// <param name="font">The font containing the glyph</param>
        /// <param name="glyphId">The glyph ID in that font</param>
        /// <returns>True if a font was found that contains the glyph</returns>
        bool TryGetGlyphFont(uint codePoint, out OpenTypeFont font, out ushort glyphId);

        /// <summary>
        /// Gets all fonts in the provider (primary + fallbacks).
        /// Used for subsetting and PDF embedding.
        /// </summary>
        IEnumerable<OpenTypeFont> GetAllFonts();
    }
}