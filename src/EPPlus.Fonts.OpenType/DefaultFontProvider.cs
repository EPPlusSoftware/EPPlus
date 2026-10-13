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
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontCache;
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Logging;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
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
    ///
    /// The decisions are reported to the logger configured on the font source, if any. A glyph
    /// found in the primary font is never logged, as that is the hot path. Everything else is
    /// reported once: the resolution of a script's chain, the first use of each fallback font,
    /// each code point routed to a fallback (debug level) and each code point no font can supply.
    /// </summary>
    public class DefaultFontProvider : IFontProvider
    {
        // Upper bound on distinct missing code points reported by one provider, so a document
        // full of unsupported characters cannot flood the log.
        private const int MaxReportedMissing = 100;

        private readonly OpenTypeFont _primaryFont;

        private readonly IFontSource _fontSource;

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

        // Log de-duplication. Allocated on first use, and only when a logger asks for the events,
        // so a provider without a logger pays nothing. Guarded by _lock.
        private HashSet<uint> _reportedMissing;
        private HashSet<uint> _reportedFallbackGlyphs;
        private bool _missingCapReported;

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
            : this(engine == null ? null : engine.FontStore, primaryFont)
        {
            // The initializer runs first, so the null guard has to be in the expression above.
            // This check exists only so the exception names 'engine' rather than 'fontSource' —
            // the caller passed an engine and should be told about an engine.
            if (engine == null)
                throw new ArgumentNullException("engine");
        }

        /// <summary>
        /// Creates a font provider over a font source. Used by the engine, which passes its own
        /// store rather than itself — a glyph provider has no business reaching a shaper factory.
        /// </summary>
        internal DefaultFontProvider(IFontSource fontSource, OpenTypeFont primaryFont)
        {
            if (fontSource == null)
                throw new ArgumentNullException("fontSource");
            if (primaryFont == null)
                throw new ArgumentNullException("primaryFont");

            _fontSource = fontSource;
            _primaryFont = primaryFont;
            _notoEmoji = new LazyFallbackFont(EmbeddedFonts.LoadNotoEmoji);
            _notoMath = new LazyFallbackFont(EmbeddedFonts.LoadNotoMath);
        }

        private IFontLogger Logger
        {
            get { return _fontSource.Logger; }
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
                    if (TryGlyphInLazyFallback(_notoEmoji, script, codePoint, out font, out glyphId))
                        return true;
                    break;

                case UnicodeScript.Math:
                    if (TryGlyphInLazyFallback(_notoMath, script, codePoint, out font, out glyphId))
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
            LogGlyphMissing(codePoint, script);
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
        /// 
        private bool TryGlyphInLazyFallback(
            LazyFallbackFont lazy,
            UnicodeScript script,
            uint codePoint,
            out OpenTypeFont font,
            out ushort glyphId)
        {
            var fallbackFont = lazy.Font; // triggers load on first use (thread-safe inside)
            if (fallbackFont.CmapTable.TryGetGlyphId(codePoint, out glyphId))
            {
                font = fallbackFont;
                MarkUsed(fallbackFont, script, codePoint);
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
                    MarkUsed(candidate, script, codePoint);
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
            List<OpenTypeFont> resolved;
            var events = new List<FontLogEvent>();

            lock (_lock)
            {
                if (_resolvedScriptChains.TryGetValue(script, out resolved))
                    return resolved;

                resolved = ResolveScriptChain(script, events);
                _resolvedScriptChains[script] = resolved;
            }

            // Reported after the lock is released, so a slow logger cannot block other threads
            // that are shaping text with this provider.
            var logger = Logger;
            foreach (var logEvent in events)
            {
                FontLog.Write(logger, logEvent);
            }

            return resolved;
        }

        /// <summary>
        /// Reads the configured chain for a script from the engine and loads each named font.
        /// Events describing the outcome are added to <paramref name="events"/> for the caller
        /// to report; nothing is logged from here, as the caller holds a lock.
        /// </summary>
        private List<OpenTypeFont> ResolveScriptChain(UnicodeScript script, List<FontLogEvent> events)
        {
            var result = new List<OpenTypeFont>();

            var logger = Logger;
            var wantInformation = FontLog.IsEnabled(logger, FontLogSeverity.Information);
            var wantWarning = FontLog.IsEnabled(logger, FontLogSeverity.Warning);
            var primary = wantInformation || wantWarning ? FontLog.Describe(_primaryFont) : null;

            var chainNames = _fontSource.GetScriptFallback(script);
            if (chainNames == null || chainNames.Length == 0)
            {
                if (wantInformation)
                {
                    events.Add(new FontLogEvent
                    {
                        Type = FontLogEventType.ScriptChainResolved,
                        Severity = FontLogSeverity.Information,
                        Script = script,
                        RequestedFont = primary,
                        Message = string.Format(
                            "Script chain {0} (primary {1}): {2}.",
                            script, primary,
                            chainNames == null ? "no chain configured" : "fallback disabled (empty chain)")
                    });
                }
                return result;
            }

            var status = wantInformation ? new List<string>() : null;

            foreach (var fontName in chainNames)
            {
                if (string.IsNullOrEmpty(fontName))
                    continue;

                // Only accept exact matches — falling back from "Microsoft YaHei" to Archivo
                // Narrow defeats the purpose of script fallback. We rely on the engine's
                // availability check rather than blindly loading.
                var availability = _fontSource.GetFontAvailability(fontName, FontSubFamily.Regular);
                if (availability != FontAvailability.Exact)
                {
                    if (status != null)
                    {
                        status.Add(fontName + (availability == FontAvailability.FamilyOnly
                            ? " [family only, not exact]"
                            : " [not found]"));
                    }
                    continue;
                }

                try
                {
                    var font = _fontSource.LoadFont(fontName, FontSubFamily.Regular);
                    if (font != null)
                    {
                        result.Add(font);
                        if (status != null)
                            status.Add(fontName + " [ok]");
                    }
                    else if (status != null)
                    {
                        status.Add(fontName + " [load returned null]");
                    }
                }
                catch (Exception ex)
                {
                    // If a named fallback fails to load for any reason, skip it.
                    // The chain is best-effort — we never want a fallback font's loading
                    // error to break primary text rendering. The failure is reported, though.
                    if (status != null)
                        status.Add(fontName + " [load failed]");

                    if (wantWarning)
                    {
                        events.Add(new FontLogEvent
                        {
                            Type = FontLogEventType.ScriptFontLoadFailed,
                            Severity = FontLogSeverity.Warning,
                            Script = script,
                            RequestedFont = fontName,
                            Exception = ex,
                            Message = string.Format(
                                "Script chain {0}: font '{1}' is installed but could not be loaded ({2}: {3}).",
                                script, fontName, ex.GetType().Name, ex.Message)
                        });
                    }
                }
            }

            if (wantInformation)
            {
                events.Add(new FontLogEvent
                {
                    Type = FontLogEventType.ScriptChainResolved,
                    Severity = FontLogSeverity.Information,
                    Script = script,
                    RequestedFont = primary,
                    Message = string.Format(
                        "Script chain {0} (primary {1}): {2}.",
                        script, primary, string.Join(", ", status.ToArray()))
                });
            }

            return result;
        }

        /// <summary>
        /// Records that a fallback font supplied a glyph. The first time a font is used, and the
        /// first time a code point is routed to a fallback, are reported.
        /// </summary>
        private void MarkUsed(OpenTypeFont font, UnicodeScript script, uint codePoint)
        {
            bool firstUseOfFont;
            lock (_lock)
            {
                firstUseOfFont = _usedFallbacks.Add(font);
            }

            var logger = Logger;

            if (firstUseOfFont && FontLog.IsEnabled(logger, FontLogSeverity.Information))
            {
                var primary = FontLog.Describe(_primaryFont);
                var used = FontLog.Describe(font);
                FontLog.Write(logger, new FontLogEvent
                {
                    Type = FontLogEventType.ScriptFallbackUsed,
                    Severity = FontLogSeverity.Information,
                    Script = script,
                    CodePoint = codePoint,
                    RequestedFont = primary,
                    ResolvedFont = used,
                    Message = string.Format(
                        "Glyph fallback: {0} lacks {1} ({2}); using {3}.",
                        primary, FontLog.FormatCodePoint(codePoint), script, used)
                });
            }

            if (FontLog.IsEnabled(logger, FontLogSeverity.Debug))
            {
                bool firstUseOfCodePoint;
                lock (_lock)
                {
                    if (_reportedFallbackGlyphs == null)
                        _reportedFallbackGlyphs = new HashSet<uint>();
                    firstUseOfCodePoint = _reportedFallbackGlyphs.Add(codePoint);
                }

                if (firstUseOfCodePoint)
                {
                    var primary = FontLog.Describe(_primaryFont);
                    var used = FontLog.Describe(font);
                    FontLog.Write(logger, new FontLogEvent
                    {
                        Type = FontLogEventType.GlyphFallback,
                        Severity = FontLogSeverity.Debug,
                        Script = script,
                        CodePoint = codePoint,
                        RequestedFont = primary,
                        ResolvedFont = used,
                        Message = string.Format(
                            "{0} ({1}) -> {2}.",
                            FontLog.FormatCodePoint(codePoint), script, used)
                    });
                }
            }
        }

        /// <summary>
        /// Reports a code point that no candidate font could supply. Each code point is reported
        /// once per provider, up to <see cref="MaxReportedMissing"/> distinct code points.
        /// </summary>
        private void LogGlyphMissing(uint codePoint, UnicodeScript script)
        {
            var logger = Logger;
            if (!FontLog.IsEnabled(logger, FontLogSeverity.Warning))
                return;

            var capReached = false;
            lock (_lock)
            {
                if (_reportedMissing == null)
                    _reportedMissing = new HashSet<uint>();

                if (_reportedMissing.Contains(codePoint))
                    return;

                if (_reportedMissing.Count >= MaxReportedMissing)
                {
                    if (_missingCapReported)
                        return;
                    _missingCapReported = true;
                    capReached = true;
                }
                else
                {
                    _reportedMissing.Add(codePoint);
                }
            }

            var primary = FontLog.Describe(_primaryFont);

            if (capReached)
            {
                FontLog.Write(
                    logger,
                    FontLogSeverity.Warning,
                    FontLogEventType.GlyphMissing,
                    string.Format(
                        "More than {0} distinct glyphs are missing from {1} and its fallbacks; further ones are not reported.",
                        MaxReportedMissing, primary),
                    primary,
                    null);
                return;
            }

            FontLog.Write(logger, new FontLogEvent
            {
                Type = FontLogEventType.GlyphMissing,
                Severity = FontLogSeverity.Warning,
                Script = script,
                CodePoint = codePoint,
                RequestedFont = primary,
                Message = string.Format(
                    "Glyph missing: {0} ({1}) is not in {2}; {3}.",
                    FontLog.FormatCodePoint(codePoint), script, primary, DescribeWhyMissing(script))
            });
        }

        /// <summary>
        /// Explains why a code point of the given script ended up without a glyph. Called only
        /// when a logger wants the event, and never from inside the lock.
        /// </summary>
        private string DescribeWhyMissing(UnicodeScript script)
        {
            switch (script)
            {
                case UnicodeScript.Unknown:
                    return "the code point has no script classification, so no fallback applies";

                case UnicodeScript.Emoji:
                    return "the bundled Noto Emoji has no glyph for it either";

                case UnicodeScript.Math:
                    return "the bundled Noto Math has no glyph for it either";
            }

            var configured = _fontSource.GetScriptFallback(script);
            if (configured == null)
                return "no fallback chain is configured for this script";
            if (configured.Length == 0)
                return "fallback is disabled for this script";

            var resolved = GetOrResolveScriptChain(script);
            if (resolved.Count == 0)
            {
                return "no font in the chain (" + FontLog.JoinNames(configured)
                    + ") is installed and loadable";
            }

            var names = new List<string>();
            foreach (var f in resolved)
            {
                names.Add(FontLog.Describe(f));
            }
            return "none of the loaded chain fonts (" + string.Join(", ", names.ToArray()) + ") has it";
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