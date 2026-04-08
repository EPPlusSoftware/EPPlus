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
  02/24/2026         EPPlus Software AB           Dynamic fallback chain with lazy loading
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Default font provider with automatic embedded fallback fonts.
    /// Fallback fonts are lazy-loaded on first use (thread-safe).
    /// Default chain: Primary → Noto Emoji → Noto Sans Math.
    /// </summary>
    public class DefaultFontProvider : IFontProvider
    {
        private readonly OpenTypeFont _primaryFont;
        private readonly List<LazyFallbackFont> _fallbackFonts;
        private readonly object _lock = new object();

        public OpenTypeFont PrimaryFont
        {
            get { return _primaryFont; }
        }

        /// <summary>
        /// Creates a font provider with automatic embedded fallbacks.
        /// Default fallback chain: Noto Emoji → Noto Sans Math.
        /// </summary>
        /// <param name="primaryFont">The user's primary font</param>
        public DefaultFontProvider(OpenTypeFont primaryFont)
        {
            if (primaryFont == null)
                throw new ArgumentNullException("primaryFont");

            _primaryFont = primaryFont;
            _fallbackFonts = new List<LazyFallbackFont>
            {
                new LazyFallbackFont(EmbeddedFonts.LoadNotoEmoji),
                new LazyFallbackFont(EmbeddedFonts.LoadNotoMath)
            };
        }

        public bool TryGetGlyphFont(uint codePoint, out OpenTypeFont font, out ushort glyphId)
        {
            // Try primary font first
            if (_primaryFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = _primaryFont;
                return true;
            }

            // Try fallback fonts in order (lazy-loaded)
            lock (_lock)
            {
                foreach (var fallback in _fallbackFonts)
                {
                    var fallbackFont = fallback.Font;
                    if (fallbackFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
                    {
                        font = fallbackFont;
                        return true;
                    }
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

            lock (_lock)
            {
                foreach (var fallback in _fallbackFonts)
                {
                    if (fallback.IsLoaded)
                    {
                        yield return fallback.Font;
                    }
                }
            }
        }

        /// <summary>
        /// Wraps a font loader delegate with lazy, thread-safe initialization.
        /// </summary>
        private class LazyFallbackFont
        {
            private readonly Func<OpenTypeFont> _loader;
            private OpenTypeFont _font;
            private readonly object _lock = new object();

            internal LazyFallbackFont(Func<OpenTypeFont> loader)
            {
                _loader = loader;
            }

            /// <summary>
            /// Gets whether the font has been loaded yet.
            /// </summary>
            internal bool IsLoaded
            {
                get { return _font != null; }
            }

            /// <summary>
            /// Gets the font, loading it on first access.
            /// </summary>
            internal OpenTypeFont Font
            {
                get
                {
                    if (_font == null)
                    {
                        lock (_lock)
                        {
                            if (_font == null)
                            {
                                _font = _loader();
                            }
                        }
                    }
                    return _font;
                }
            }
        }
    }
}