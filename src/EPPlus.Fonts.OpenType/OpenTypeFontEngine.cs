/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/13/2026         EPPlus Software AB           Per-instance font engine. Replaces static OpenTypeFonts.
  09/02/2026         EPPlus Software AB           Extracted FontStore and ShaperCache; added measurement shaper
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontCache;
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// An OpenType font engine instance. Produces shapers and layout engines, and owns the
    /// configuration, the font store and the shaper cache for its own lifetime. Two engines do
    /// not share parsed fonts and can have different configurations simultaneously.
    ///
    /// Configuration is set at construction time and is immutable for the lifetime of the engine.
    /// To use a different configuration, create a new engine.
    ///
    /// The engine holds the policy for which kind of shaper a request gets; the store and the
    /// cache hold no policy at all.
    /// </summary>
    public class OpenTypeFontEngine : IDisposable
    {
        private readonly EpplusFontConfiguration _configuration;
        private readonly FontStore _fontStore;
        private readonly ShaperCache _shaperCache = new ShaperCache();

        private bool _disposed;

        /// <summary>
        /// Creates an engine with default configuration: no extra font directories,
        /// system directories searched, no user fallbacks, default resolver.
        /// </summary>
        public OpenTypeFontEngine()
            : this(null)
        {
        }

        /// <summary>
        /// Creates an engine with the given configuration.
        /// The configure callback runs once during construction and mutates the new
        /// configuration. After the constructor returns, the configuration is fixed.
        /// </summary>
        /// <example>
        /// var engine = new OpenTypeFontEngine(cfg =>
        /// {
        ///     cfg.FontDirectories.Add(@"C:\MyApp\Fonts");
        ///     cfg.SearchSystemDirectories = false;
        ///     cfg.FontFallbacks["Arial"] = new[] { "Helvetica" };
        /// });
        /// </example>
        public OpenTypeFontEngine(Action<IEpplusFontConfiguration> configure)
        {
            _configuration = new EpplusFontConfiguration();

            if (configure != null)
            {
                configure(_configuration);
            }

            // If the user installed a custom resolver, use it as-is. Otherwise build a
            // DefaultFontResolver from the configuration.
            var resolver = _configuration.FontResolver;
            if (resolver == null)
            {
                resolver = new DefaultFontResolver(
                    fontDirectories: _configuration.FontDirectories,
                    searchSystemDirectories: _configuration.SearchSystemDirectories,
                    config: _configuration);
            }

            _fontStore = new FontStore(resolver, _configuration);
        }

        // -----------------------------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// When true, shaper resolution throws if the requested font cannot be resolved to an
        /// exact match, even though a fallback was found. Default is false: rendering trusts the
        /// fallback chain (which always resolves to at least the embedded font) and never throws.
        /// Set to true only for diagnostics or validation where a missing exact font should
        /// surface as an error.
        ///
        /// It also suppresses the metrics fallback on the measurement path, since silently
        /// substituting serialized metrics would defeat the purpose of the mode.
        /// </summary>
        public bool RequireExactFont { get; set; } = false;

        /// <summary>
        /// The font store backing this engine. Internal: it is how the engine hands an
        /// <see cref="IFontSource"/> to the providers it constructs, and how the public
        /// <see cref="DefaultFontProvider"/> constructor forwards to its internal one.
        /// </summary>
        internal FontStore FontStore
        {
            get { return _fontStore; }
        }

        // -----------------------------------------------------------------------------------------
        // Shapers — this is the policy
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Gets a <see cref="TextShaper"/> for the given font, reusing a thread-local cached
        /// instance. The underlying OpenTypeFont is shared within this engine (but not between
        /// engines), while each thread gets its own TextShaper to avoid locking.
        /// Returns null if the font cannot be resolved.
        ///
        /// The returned shaper is always backed by a real font file, so its output carries glyph
        /// ids, glyph outlines and font references. Use this when the caller needs glyph data:
        /// PDF export, subsetting, embedding. For measurement and line breaking use
        /// <see cref="GetMeasurementShaper"/> instead.
        /// </summary>
        public TextShaper GetTextShaper(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            ThrowIfDisposed();
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            return GetOrCreateRenderingShaper(FontStore.BuildCacheKey(fontName, subFamily), fontName, subFamily);
        }

        /// <summary>
        /// Gets a shaper for measuring text and breaking lines.
        ///
        /// Unlike <see cref="GetTextShaper"/> this may return a shaper that is not backed by a
        /// font file. When the requested font cannot be resolved and font resolution has fallen
        /// all the way through to the embedded last-resort font, the serialized font metrics
        /// (.fmtr) for the requested family are used instead, when they exist. Measuring the
        /// requested family from quantized metrics is closer to the truth than measuring a
        /// different family exactly.
        ///
        /// The returned shaper may have no glyph ids, no kerning and no OpenType layout tables;
        /// see <see cref="ITextShaper.HasGlyphIds"/>. The narrower return type is deliberate — it
        /// exposes no glyph-level API, so a glyph consumer cannot reach for this method by
        /// accident.
        ///
        /// Returns null only when neither a font file nor serialized metrics can be found, which
        /// requires a custom <see cref="IFontResolver"/> that returns null.
        /// </summary>
        public ITextShaper GetMeasurementShaper(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            ThrowIfDisposed();
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            var key = FontStore.BuildCacheKey(fontName, subFamily);

            ITextShaper cached;
            if (_shaperCache.TryGetMeasurement(key, out cached))
            {
                return cached;
            }

            var shaper = CreateMeasurementShaper(key, fontName, subFamily);

            // A null result is not cached. It only happens with a custom resolver that gives up,
            // and caching it would make a later change in resolver state unobservable.
            if (shaper != null)
            {
                _shaperCache.AddMeasurement(key, shaper);
            }

            return shaper;
        }

        /// <summary>
        /// Gets a shaper for measuring text in the given font. May return a metrics-only shaper;
        /// see <see cref="GetMeasurementShaper"/>.
        /// </summary>
        public ITextShaper GetShaperForFont(IFontFormatBase font)
        {
            if (font == null)
                throw new ArgumentNullException("font");

            return GetMeasurementShaper(font.Family, font.SubFamily);
        }

        /// <summary>
        /// Gets a shaper for measuring text in the given font. May return a metrics-only shaper;
        /// see <see cref="GetMeasurementShaper"/>.
        /// </summary>
        public ITextShaper GetShaperForFont(MeasurementFont font)
        {
            if (font == null)
                throw new ArgumentNullException("font");

            return GetMeasurementShaper(font.FontFamily, FontSubFamilyConverter.ToSubFamily(font.Style));
        }

        // -----------------------------------------------------------------------------------------
        // Layout engines
        //
        // A TextLayoutEngine only ever measures, so all three of these take the measurement path.
        // -----------------------------------------------------------------------------------------

        public TextLayoutEngine GetTextLayoutEngine(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            var shaper = GetMeasurementShaper(fontName, subFamily);
            return new TextLayoutEngine(this, shaper);
        }

        public TextLayoutEngine GetTextLayoutEngineForFont(IFontFormatBase font)
        {
            var shaper = GetShaperForFont(font);
            return new TextLayoutEngine(this, shaper);
        }

        public TextLayoutEngine GetTextLayoutEngineForFont(MeasurementFont font)
        {
            var shaper = GetShaperForFont(font);
            return new TextLayoutEngine(this, shaper);
        }

        // -----------------------------------------------------------------------------------------
        // Font queries — delegating to the store
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Loads a font by name and subfamily, with thread-safe caching within this engine.
        /// Returns null if the font cannot be resolved.
        /// </summary>
        public OpenTypeFont LoadFont(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            bool ignoreCache = false)
        {
            ThrowIfDisposed();
            return _fontStore.LoadFont(fontName, subFamily, ignoreCache);
        }

        /// <summary>
        /// Checks whether a font is available in this engine's configured font system.
        /// See <see cref="FontStore.GetFontAvailability"/> for the accuracy caveats.
        /// </summary>
        public FontAvailability GetFontAvailability(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular)
        {
            ThrowIfDisposed();
            return _fontStore.GetFontAvailability(fontName, subFamily);
        }

        /// <summary>
        /// Creates an OpenTypeFont directly from raw font bytes.
        /// Font format (TTF/OTF) is detected automatically from the SFNT header.
        /// Independent of engine configuration — consults neither the resolver nor the cache.
        /// </summary>
        public OpenTypeFont GetFromBytes(byte[] bytes)
        {
            if (bytes == null)
                throw new ArgumentNullException("bytes");

            var font = new OpenTypeFont(bytes);
            font.EnsureFullyLoaded();
            return font;
        }


        // -----------------------------------------------------------------------------------------
        // Lifecycle
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Clears this engine's parsed-font cache, per-font locks, and the calling thread's
        /// shaper cache for this engine. Does not affect scanner-level caches, which are global
        /// and reflect filesystem state rather than engine configuration.
        /// </summary>
        public void ClearFontCache()
        {
            ThrowIfDisposed();

            // Shapers first. Each retained shaper holds a reference to a parsed font, so
            // clearing the fonts first would leave live shapers pointing at fonts that are no
            // longer in the cache and would be re-parsed on the next miss.
            _shaperCache.ClearCurrentThread();
            _fontStore.Clear();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            // Best-effort cleanup of this engine's shaper entries on the disposing thread.
            // Entries on other threads are dropped when those threads next touch their map,
            // since they will find no entry for this cache and rebuild.
            _shaperCache.ClearCurrentThread();

            // Marks the store disposed as well as clearing it, so a DefaultFontProvider that
            // holds the store directly cannot keep loading fonts after this point.
            _fontStore.MarkDisposed();
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException("OpenTypeFontEngine");
        }

        internal FontEmbeddingDecision ResolveEmbeddingDecision(OpenTypeFont font)
        {
            var restriction = font.Os2Table != null
                ? font.Os2Table.GetEmbeddingRestriction()
                : FontEmbeddingRestriction.None;

            if (font.NameTable != null && EmbeddedFonts.IsBundledFamily(font.GetEnglishFontFamilyName()))
                return FontEmbeddingDecision.Subset;

            var fontName = font.NameTable != null ? font.NameTable.GetFullFontName() : null;
            var callback = _configuration.GetEmbeddingCallback();
            if (callback != null)
            {
                var decision = callback(new FontEmbeddingInfo(fontName, restriction));
                if (decision != FontEmbeddingDecision.Default)
                    return decision;   // user override wins
            }

            Debug.WriteLine($"ResolveEmbeddingDecision: {fontName} restriction={restriction} callback={(callback != null)}");

            // No callback, or callback returned Default → derive from the restriction.
            switch (restriction)
            {
                case FontEmbeddingRestriction.NoEmbedding:
                    // Default policy: fail loud. User must opt in via the callback.
                    throw new InvalidOperationException(
                        string.Format(
                            "Font '{0}' declares Restricted License embedding (fsType) and may not be embedded. " +
                            "If you hold a licence permitting embedding, return FontEmbeddingDecision.Subset or " +
                            "EmbedWhole from IEpplusFontConfiguration.OnFontEmbedding.",
                            string.IsNullOrWhiteSpace(fontName) ? "(unknown)" : fontName));
                case FontEmbeddingRestriction.NoSubsetting:
                    return FontEmbeddingDecision.EmbedWhole;
                default:
                    return FontEmbeddingDecision.Subset;
            }
        }

        // -----------------------------------------------------------------------------------------
        // Shaper resolution — policy
        // -----------------------------------------------------------------------------------------

        private TextShaper GetOrCreateRenderingShaper(string key, string fontName, FontSubFamily subFamily)
        {
            TextShaper shaper;
            if (_shaperCache.TryGetRendering(key, out shaper))
            {
                return shaper;
            }

            var font = _fontStore.LoadFont(fontName, subFamily, false);
            if (font == null)
                return null;

            ThrowIfNotExactWhenRequired(fontName, subFamily, font);

            shaper = CreateShaper(font);
            _shaperCache.AddRendering(key, shaper);
            return shaper;
        }

        private ITextShaper CreateMeasurementShaper(string key, string fontName, FontSubFamily subFamily)
        {
            var fallbackMode = _configuration.MetricsFallback;

            // Always: short-circuit before LoadFont so no font file is opened at all. This is
            // what makes measurement reproducible across machines — the result cannot depend on
            // what happens to be installed. RequireExactFont wins, since a diagnostic mode that
            // wants a missing font to throw must not be silenced by a metrics substitution.
            if (fallbackMode == MetricsFallbackMode.Always && !RequireExactFont)
            {
                GenericFontTextShaper alwaysShaper;
                if (GenericFontTextShaper.TryCreate(fontName, FontSubFamilyConverter.ToStyles(subFamily), out alwaysShaper))
                {
                    return alwaysShaper;
                }
                // No metrics for this family. Fall through to normal resolution rather than
                // failing — Always is a preference, not a constraint.
            }

            // If the rendering path already parsed this font on this thread, reuse that shaper.
            // A real font is the better measurement source and the instance is identical.
            TextShaper alreadyParsed;
            if (_shaperCache.TryGetRendering(key, out alreadyParsed))
            {
                return alreadyParsed;
            }

            var font = _fontStore.LoadFont(fontName, subFamily, false);

            if (fallbackMode != MetricsFallbackMode.Disabled
                && !RequireExactFont
                && ResolvedToLastResort(fontName, font))
            {
                GenericFontTextShaper metricsShaper;
                if (GenericFontTextShaper.TryCreate(fontName, FontSubFamilyConverter.ToStyles(subFamily), out metricsShaper))
                {
                    return metricsShaper;
                }
            }

            if (font == null)
                return null;

            ThrowIfNotExactWhenRequired(fontName, subFamily, font);

            var shaper = CreateShaper(font);
            _shaperCache.AddRendering(key, shaper);
            return shaper;
        }

        /// <summary>
        /// Builds the provider and shaper for a parsed font. The provider gets the store rather
        /// than this engine, so nothing the engine creates holds a reference back to it.
        /// </summary>
        private TextShaper CreateShaper(OpenTypeFont font)
        {
            return new TextShaper(new DefaultFontProvider(_fontStore, font));
        }

        private void ThrowIfNotExactWhenRequired(string fontName, FontSubFamily subFamily, OpenTypeFont font)
        {
            if (!RequireExactFont)
                return;

            var availability = _fontStore.GetFontAvailability(fontName, subFamily);
            if (availability != FontAvailability.Exact)
            {
                throw new FileNotFoundException(
                    $"Could not find Font: {fontName} {subFamily}. Resolved via fallback to: {font.GetEnglishFontFamilyName()} {font.SubFamily}.");
            }
        }

        /// <summary>
        /// True when font resolution produced the embedded last-resort font for a request that did
        /// not ask for it, or produced nothing at all.
        ///
        /// Deliberately not expressed as GetFontAvailability returning NotFound. That reports only
        /// on the requested family and knows nothing about the user-configured and built-in
        /// fallback chains in DefaultFontResolver, so it reports NotFound even when a
        /// metric-compatible substitute was found and used. A real substitute font is a better
        /// measurement source than quantized metrics, so the metrics fallback must engage only
        /// after those chains are exhausted — that is, at the point resolution gives up and loads
        /// the embedded font.
        /// </summary>
        private static bool ResolvedToLastResort(string requestedFontName, OpenTypeFont resolvedFont)
        {
            // The caller asked for the last-resort font itself and got it. Not a fallback.
            if (IsLastResortFamily(requestedFontName))
                return false;

            // A custom IFontResolver returned null. Nothing was resolved at all.
            if (resolvedFont == null)
                return true;

            return IsLastResortFamily(resolvedFont.GetEnglishFontFamilyName());
        }

        private static bool IsLastResortFamily(string fontName)
        {
            return string.Equals("archivo narrow", fontName, StringComparison.OrdinalIgnoreCase);
        }
    }
}