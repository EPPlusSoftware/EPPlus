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
  05/20/2026         EPPlus Software AB           Script-classified fallback via engine reference
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Default font provider with script-classified glyph fallback.
    ///
    /// When a code point is missing from the primary font, the provider routes the lookup
    /// based on the code point's Unicode script:
    ///   * Emoji  → embedded Noto Emoji (bundled with EPPlus)
    ///   * Math   → embedded Noto Math (bundled with EPPlus)
    ///   * Other  → per-script fallback chain configured on the engine (best-effort,
    ///              resolves named fonts via the engine — works only when the named fonts
    ///              are installed)
    ///
    /// Per-script chains and their fonts are lazy-loaded the first time a code point in
    /// that script is encountered, then cached for the lifetime of this provider.
    /// </summary>
    public class DefaultFontProvider : IFontProvider
    {
        private readonly OpenTypeFontEngine _engine;
        private readonly OpenTypeFont _primaryFont;

        // Embedded fallbacks, lazy-loaded on first use.
        private readonly LazyFallbackFont _notoEmoji;
        private readonly LazyFallbackFont _notoMath;

        // Per-script named fallbacks, resolved via the engine on first use of each script.
        // Inner list is the resolved chain of fonts for that script; entries that fail to
        // resolve are omitted (so the list may be shorter than the configured chain).
        private readonly Dictionary<UnicodeScript, List<OpenTypeFont>> _resolvedScriptChains
            = new Dictionary<UnicodeScript, List<OpenTypeFont>>();

        // Tracks which fallback fonts have actually been used (returned a glyph for some
        // code point). Used by GetAllFonts to expose only the fonts that mattered, which
        // matters for subsetting and PDF embedding.
        private readonly HashSet<OpenTypeFont> _usedFallbacks = new HashSet<OpenTypeFont>();

        private readonly object _lock = new object();

        /// <inheritdoc/>
        public OpenTypeFont PrimaryFont
        {
            get { return _primaryFont; }
        }

        /// <summary>
        /// Creates a font provider that uses the given engine to resolve per-script fallback
        /// fonts on demand. Both arguments are required.
        /// </summary>
        /// <param name="engine">The engine to use for resolving named fallback fonts.</param>
        /// <param name="primaryFont">The primary font for text in the user's chosen typeface.</param>
        public DefaultFontProvider(OpenTypeFontEngine engine, OpenTypeFont primaryFont)
        {
            if (engine == null)
                throw new ArgumentNullException("engine");
            if (primaryFont == null)
                throw new ArgumentNullException("primaryFont");

            _engine = engine;
            _primaryFont = primaryFont;
            _notoEmoji = new LazyFallbackFont(EmbeddedFonts.LoadNotoEmoji);
            _notoMath = new LazyFallbackFont(EmbeddedFonts.LoadNotoMath);
        }

        /// <inheritdoc/>
        public bool TryGetGlyphFont(uint codePoint, out OpenTypeFont font, out ushort glyphId)
        {
            // 1. Primary font wins whenever it has the glyph.
            if (_primaryFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = _primaryFont;
                return true;
            }

            // 2. Classify the code point and route to the appropriate fallback.
            var script = UnicodeScriptClassifier.OfCodePoint(codePoint);

            switch (script)
            {
                case UnicodeScript.Emoji:
                    if (TryGlyphInLazyFallback(_notoEmoji, codePoint, out font, out glyphId))
                        return true;
                    break;

                case UnicodeScript.Math:
                    if (TryGlyphInLazyFallback(_notoMath, codePoint, out font, out glyphId))
                        return true;
                    break;

                case UnicodeScript.Unknown:
                    // No script classification — no useful fallback to route to.
                    break;

                default:
                    if (TryGlyphInScriptChain(script, codePoint, out font, out glyphId))
                        return true;
                    break;
            }

            // 3. Nothing found — return primary with .notdef.
            font = _primaryFont;
            glyphId = 0;
            return false;
        }

        /// <inheritdoc/>
        public IEnumerable<OpenTypeFont> GetAllFonts()
        {
            yield return _primaryFont;

            // Only return fallback fonts that have actually been used. Subsetting and PDF
            // embedding only need fonts whose glyphs the shaper actually placed.
            lock (_lock)
            {
                foreach (var f in _usedFallbacks)
                {
                    yield return f;
                }
            }
        }

        // -----------------------------------------------------------------------------------------
        // Internal helpers
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Tries to find the glyph in a lazy-loaded embedded fallback font (Noto Emoji / Math).
        /// </summary>
        private bool TryGlyphInLazyFallback(
            LazyFallbackFont lazy,
            uint codePoint,
            out OpenTypeFont font,
            out ushort glyphId)
        {
            var fallbackFont = lazy.Font; // triggers load on first use (thread-safe inside)
            if (fallbackFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = fallbackFont;
                MarkUsed(fallbackFont);
                return true;
            }

            font = null;
            glyphId = 0;
            return false;
        }

        /// <summary>
        /// Tries to find the glyph by walking the per-script fallback chain configured on
        /// the engine. Resolves the chain lazily on first use of each script.
        /// </summary>
        private bool TryGlyphInScriptChain(
            UnicodeScript script,
            uint codePoint,
            out OpenTypeFont font,
            out ushort glyphId)
        {
            var chain = GetOrResolveScriptChain(script);

            foreach (var candidate in chain)
            {
                if (candidate.CmapTable.TryGetGlyphId(codePoint, out glyphId))
                {
                    font = candidate;
                    MarkUsed(candidate);
                    return true;
                }
            }

            font = null;
            glyphId = 0;
            return false;
        }

        /// <summary>
        /// Returns the resolved chain of fonts for a script. The first time a script is
        /// queried, the configured chain of font names is read from the engine's configuration
        /// and each name is resolved via the engine. Names that fail to resolve are omitted.
        /// </summary>
        private List<OpenTypeFont> GetOrResolveScriptChain(UnicodeScript script)
        {
            lock (_lock)
            {
                List<OpenTypeFont> resolved;
                if (_resolvedScriptChains.TryGetValue(script, out resolved))
                    return resolved;

                resolved = ResolveScriptChain(script);
                _resolvedScriptChains[script] = resolved;
                return resolved;
            }
        }

        /// <summary>
        /// Reads the configured chain for a script from the engine and loads each named font.
        /// </summary>
        private List<OpenTypeFont> ResolveScriptChain(UnicodeScript script)
        {
            var result = new List<OpenTypeFont>();

            var chainNames = _engine.GetScriptFallback(script);
            if (chainNames == null || chainNames.Length == 0)
                return result;

            foreach (var fontName in chainNames)
            {
                if (string.IsNullOrEmpty(fontName))
                    continue;

                // Only accept exact matches — falling back from "Microsoft YaHei" to Archivo
                // Narrow defeats the purpose of script fallback. We rely on the engine's
                // availability check rather than blindly loading.
                var availability = _engine.GetFontAvailability(fontName, FontSubFamily.Regular);
                if (availability != FontAvailability.Exact)
                    continue;

                try
                {
                    var font = _engine.LoadFont(fontName, FontSubFamily.Regular);
                    if (font != null)
                        result.Add(font);
                }
                catch
                {
                    // If a named fallback fails to load for any reason, skip it silently.
                    // The chain is best-effort — we never want a fallback font's loading
                    // error to break primary text rendering.
                }
            }

            return result;
        }

        private void MarkUsed(OpenTypeFont font)
        {
            lock (_lock)
            {
                _usedFallbacks.Add(font);
            }
        }

        /// <summary>
        /// Wraps an embedded font loader with lazy, thread-safe initialization.
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