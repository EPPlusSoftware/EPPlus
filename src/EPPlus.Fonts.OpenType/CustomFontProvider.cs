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
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Custom font provider where user defines their own fallback chain.
    /// Does NOT use embedded Noto Emoji - user must add fallbacks manually.
    /// </summary>
    public class CustomFontProvider : IFontProvider
    {
        private readonly OpenTypeFont _primaryFont;
        private readonly List<OpenTypeFont> _fallbackFonts;

        public OpenTypeFont PrimaryFont
        {
            get { return _primaryFont; }
        }

        /// <summary>
        /// Creates a custom font provider without any fallbacks.
        /// </summary>
        /// <param name="primaryFont">The user's primary font</param>
        public CustomFontProvider(OpenTypeFont primaryFont)
        {
            if (primaryFont == null)
                throw new ArgumentNullException("primaryFont");

            _primaryFont = primaryFont;
            _fallbackFonts = new List<OpenTypeFont>();
        }

        /// <summary>
        /// Adds a fallback font to the chain.
        /// Fonts are searched in the order they are added.
        /// </summary>
        /// <param name="font">Fallback font to add</param>
        public void AddFallback(OpenTypeFont font)
        {
            if (font == null)
                throw new ArgumentNullException("font");

            _fallbackFonts.Add(font);
        }

        public bool TryGetGlyphFont(uint codePoint, out OpenTypeFont font, out ushort glyphId)
        {
            // Try primary font first
            if (_primaryFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = _primaryFont;
                return true;
            }

            // Try fallback fonts in order
            foreach (var fallbackFont in _fallbackFonts)
            {
                if (fallbackFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
                {
                    font = fallbackFont;
                    return true;
                }
            }

            // Not found - return primary with .notdef
            font = _primaryFont;
            glyphId = 0;
            return false;
        }

        public IEnumerable<OpenTypeFont> GetAllFonts()
        {
            yield return _primaryFont;

            foreach (var fallback in _fallbackFonts)
            {
                yield return fallback;
            }
        }
    }
}