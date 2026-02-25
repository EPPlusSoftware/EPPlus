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
    /// Default font provider with automatic Noto Emoji fallback.
    /// </summary>
    public class DefaultFontProvider : IFontProvider
    {
        private readonly OpenTypeFont _primaryFont;
        private OpenTypeFont _emojiFont;
        private readonly object _lock = new object();

        public OpenTypeFont PrimaryFont
        {
            get { return _primaryFont; }
        }

        /// <summary>
        /// Creates a font provider with automatic emoji fallback.
        /// </summary>
        /// <param name="primaryFont">The user's primary font</param>
        public DefaultFontProvider(OpenTypeFont primaryFont)
        {
            if (primaryFont == null)
                throw new ArgumentNullException("primaryFont");

            _primaryFont = primaryFont;
        }

        /// <summary>
        /// Gets the emoji font, loading it on first access (lazy loading).
        /// Thread-safe for .NET 3.5 compatibility.
        /// </summary>
        private OpenTypeFont GetEmojiFontLazy()
        {
            if (_emojiFont == null)
            {
                lock (_lock)
                {
                    if (_emojiFont == null)
                    {
                        _emojiFont = EmbeddedFonts.LoadNotoEmoji();
                    }
                }
            }
            return _emojiFont;
        }

        public bool TryGetGlyphFont(uint codePoint, out OpenTypeFont font, out ushort glyphId)
        {
            // Try primary font first
            if (_primaryFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = _primaryFont;
                return true;
            }

            // Fallback to embedded Noto Emoji (lazy-loaded)
            var emojiFont = GetEmojiFontLazy();
            if (emojiFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = emojiFont;
                return true;
            }

            // Not found - return primary with .notdef
            font = _primaryFont;
            glyphId = 0;
            return false;
        }

        public IEnumerable<OpenTypeFont> GetAllFonts()
        {
            yield return _primaryFont;

            // Only include emoji font if it was actually loaded
            if (_emojiFont != null)
            {
                yield return _emojiFont;
            }
        }
    }
}
