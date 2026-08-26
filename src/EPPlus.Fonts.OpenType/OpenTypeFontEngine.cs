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
    /// An OpenType font engine instance. Owns its own configuration, resolver, and font cache —
    /// two engines do not share parsed fonts and can have different configurations simultaneously.
    /// Scanner-level data (file system listings, per-file FontFaceInfo) is shared globally
    /// because that data describes the filesystem and is identical regardless of engine.
    /// Configuration is set at construction time and is immutable for the lifetime of the engine.
    /// To use a different configuration, create a new engine.
    /// </summary>
    public class OpenTypeFontEngine : IDisposable
    {
        private readonly object _syncRoot = new object();
        private readonly Dictionary<string, object> _fontLocks = new Dictionary<string, object>();

        // Active resolver. Set at construction; never replaced.
        private readonly IFontResolver _fontResolver;

        // Configuration snapshot. Held to support GetFontAvailability and similar queries.
        private readonly EpplusFontConfiguration _configuration;

        // Per-engine cache of parsed fonts.
        private readonly OpenTypeFontCache _fontCache = new OpenTypeFontCache();

        // Thread-local TextShaper cache. Each thread gets its own dictionary, keyed by the engine
        // instance to avoid collisions if multiple engines are used on the same thread.
        // We use a ThreadStatic Dictionary<OpenTypeFontEngine, Dictionary<string, TextShaper>>
        // so each engine has its own per-thread shaper namespace.
        [ThreadStatic]
        private static Dictionary<OpenTypeFontEngine, Dictionary<string, TextShaper>> _threadLocalShaperCaches;

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
            var userResolver = _configuration.FontResolver;
            if (userResolver != null)
            {
                _fontResolver = userResolver;
            }
            else
            {
                _fontResolver = new DefaultFontResolver(
                    fontDirectories: _configuration.FontDirectories,
                    searchSystemDirectories: _configuration.SearchSystemDirectories,
                    config: _configuration);
            }
        }

        // -----------------------------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// When true, GetTextShaper throws if the requested font cannot be resolved to an exact match,
        /// even though a fallback was found. Default is false: rendering trusts the fallback chain
        /// (which always resolves to at least the embedded font) and never throws. Set to true only
        /// for diagnostics or validation where a missing exact font should surface as an error.
        /// </summary>
        public bool RequireExactFont { get; set; } = false;
        //public FontAvailability FallBackAvailablility = FontAvailability.Exact;

        /// <summary>
        /// Gets a TextShaper for the given font, reusing a thread-local cached instance.
        /// The underlying OpenTypeFont is shared within this engine (but not between engines),
        /// while each thread gets its own TextShaper instance to avoid locking.
        /// Returns null if the font cannot be resolved.
        /// </summary>
        public TextShaper GetTextShaper(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            ThrowIfDisposed();
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            var perEngineMap = GetOrCreateThreadLocalShaperMap();

            string key = BuildCacheKey(fontName, subFamily);

            TextShaper shaper;
            if (!perEngineMap.TryGetValue(key, out shaper))
            {
                var font = LoadFont(fontName, subFamily);
                if (font == null)
                    return null;

                if (RequireExactFont)
                {
                    var availability = GetFontAvailability(fontName, subFamily);
                    if (availability != FontAvailability.Exact)
                    {
                        throw new FileNotFoundException(
                            $"Could not find Font: {fontName} {subFamily}. Resolved via fallback to: {font.GetEnglishFontFamilyName()} {font.SubFamily}.");
                    }
                }

                shaper = new TextShaper(this, font);
                perEngineMap[key] = shaper;
            }

            return shaper;
        }

        public TextLayoutEngine GetTextLayoutEngine(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            var shaper = GetTextShaper(fontName, subFamily);
            return new TextLayoutEngine(this, shaper);
        }
        public TextLayoutEngine GetTextLayoutEngineForFont(IFontFormatBase font)
        {
            var shaper = GetShaperForFont(font);
            return new TextLayoutEngine(this, shaper);
        }

        public ITextShaper GetShaperForFont(IFontFormatBase font)
        {
            return GetTextShaper(font.Family, font.SubFamily);
        }

        public TextLayoutEngine GetTextLayoutEngineForFont(MeasurementFont font)
        {
            var shaper = GetShaperForFont(font);
            return new TextLayoutEngine(this, shaper);
        }

        public ITextShaper GetShaperForFont(MeasurementFont font)
        {
            return GetTextShaper(font.FontFamily, GetFontSubFamily(font.Style));
        }

        public static FontSubFamily GetFontSubFamily(MeasurementFontStyles style)
        {
            if ((style & (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic)) ==
                (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if ((style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold)
            {
                return FontSubFamily.Bold;
            }
            else if ((style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic)
            {
                return FontSubFamily.Italic;
            }

            return FontSubFamily.Regular;
        }

        /// <summary>
        /// Clears this engine's parsed-font cache, per-font locks, and the calling thread's
        /// TextShaper cache for this engine. Does not affect scanner-level caches (which are
        /// global and reflect filesystem state, not engine configuration).
        /// </summary>
        public void ClearFontCache()
        {
            ThrowIfDisposed();
            lock (_syncRoot)
            {
                _fontCache.Clear();
                _fontLocks.Clear();
            }

            // Clear this engine's shaper cache for the calling thread.
            // Other threads' caches will be lazily rebuilt on next use.
            if (_threadLocalShaperCaches != null)
                _threadLocalShaperCaches.Remove(this);
        }

        // -----------------------------------------------------------------------------------------
        // Font loading
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

            if (ignoreCache)
                return ResolveAndCreate(_fontResolver, fontName, subFamily);

            string lockKey = BuildCacheKey(fontName, subFamily);
            object fontLock;
            lock (_syncRoot)
            {
                if (!_fontLocks.TryGetValue(lockKey, out fontLock))
                {
                    fontLock = new object();
                    _fontLocks[lockKey] = fontLock;
                }
            }

            lock (fontLock)
            {
                var cached = _fontCache.GetFromCache(lockKey);
                if (cached != null && cached.Font != null && cached.IsLoaded)
                {
                    cached.Font.EnsureFullyLoaded();
                    return cached.Font;
                }

                _fontCache.BeginCache(lockKey);

                var font = ResolveAndCreate(_fontResolver, fontName, subFamily);
                if (font == null)
                    return null;

                font.EnsureFullyLoaded();
                font.IsReadOnly = true;
                _fontCache.AddToCache(font, lockKey);
                return font;
            }
        }

        /// <summary>
        /// Returns all available font faces as fully loaded OpenTypeFont instances.
        /// Skips corrupt or unreadable fonts, but logs detailed information for diagnostics.
        /// This method is NOT cached and may take significant time to complete.
        /// Note: takes fontDirectories as a parameter; this is a diagnostic / discovery API
        /// independent of the engine's configured resolver.
        /// </summary>
        public List<OpenTypeFont> GetAllBaseFontData(
            List<string> fontDirectories,
            bool searchSystemDirectories = true,
            FontFormat? formatTarget = null)
        {
            ThrowIfDisposed();

            var locations = DefaultFontLocations.GetLocationsCollection(fontDirectories, searchSystemDirectories);
            var faces = FontScannerV2.EnumerateAllFaces(locations);

            var result = new List<OpenTypeFont>(faces.Count);
            var failures = 0;

            foreach (var face in faces)
            {
                if (formatTarget.HasValue)
                {
                    string ext = Path.GetExtension(face.FilePath);
                    if (!string.IsNullOrEmpty(ext))
                    {
                        ext = ext.ToLowerInvariant();
                        var format = (ext == ".otf" || ext == ".cff")
                            ? FontFormat.Otf
                            : FontFormat.Ttf;

                        if (format != formatTarget.Value)
                            continue;
                    }
                }

                try
                {
                    var font = new OpenTypeFont(File.ReadAllBytes(face.FilePath));
                    font.EnsureFullyLoaded();
                    result.Add(font);
                }
                catch (Exception ex)
                {
                    failures++;
                    System.Diagnostics.Debug.WriteLine(
                        string.Format("[OpenTypeFontEngine] Failed to load font: {0} => {1}: {2}",
                            face.FilePath, ex.GetType().Name, ex.Message));
                }
            }

            if (failures > 0)
                System.Diagnostics.Debug.WriteLine(
                    string.Format("[OpenTypeFontEngine] {0} font(s) failed to load.", failures));

            return result;
        }

        /// <summary>
        /// Creates an OpenTypeFont directly from raw font bytes.
        /// Font format (TTF/OTF) is detected automatically from the SFNT header.
        /// Independent of engine configuration — does not consult the resolver or cache.
        /// </summary>
        public OpenTypeFont GetFromBytes(byte[] bytes)
        {
            if (bytes == null)
                throw new ArgumentNullException("bytes");

            var font = new OpenTypeFont(bytes);
            font.EnsureFullyLoaded();
            return font;
        }

        /// <summary>
        /// Checks whether a font is available in this engine's configured font system.
        /// Returns <see cref="FontAvailability.Exact"/> if the exact family and subfamily exist,
        /// <see cref="FontAvailability.FamilyOnly"/> if the family exists but not in the requested
        /// subfamily, and <see cref="FontAvailability.NotFound"/> otherwise.
        ///
        /// If the engine has a custom resolver that implements <see cref="IFontAvailabilityProvider"/>,
        /// the call delegates to the resolver. Otherwise it probes via <see cref="IFontResolver.ResolveFont"/>,
        /// which can only distinguish "found" from "not found" — never <see cref="FontAvailability.FamilyOnly"/>.
        /// </summary>
        public FontAvailability GetFontAvailability(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular)
        {
            ThrowIfDisposed();
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            var provider = _fontResolver as IFontAvailabilityProvider;
            if (provider != null)
                return provider.GetFontAvailability(fontName, subFamily);

            // Fallback: probe via ResolveFont. Cannot distinguish FamilyOnly.
            return _fontResolver.ResolveFont(fontName, subFamily) != null
                ? FontAvailability.Exact
                : FontAvailability.NotFound;
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
        // Internal helpers
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Returns the configured fallback chain for the given Unicode script, or null if
        /// none is configured. An empty array means fallback is explicitly disabled for the
        /// script. Used by DefaultFontProvider to look up script-level glyph fallbacks.
        /// </summary>
        internal string[] GetScriptFallback(UnicodeScript script)
        {
            return _configuration.GetScriptFallback(script);
        }

        internal static string BuildCacheKey(string fontName, FontSubFamily subFamily)
        {
            return string.Format("{0}_{1}", fontName, subFamily);
        }

        private static OpenTypeFont ResolveAndCreate(IFontResolver resolver, string fontName, FontSubFamily subFamily)
        {
            var bytes = resolver.ResolveFont(fontName, subFamily);
            if (bytes == null)
                return null;

            return new OpenTypeFont(bytes);
        }

        private Dictionary<string, TextShaper> GetOrCreateThreadLocalShaperMap()
        {
            // [ThreadStatic] field initializers only run on the primary thread.
            // All other threads see null and must initialize on first use.
            if (_threadLocalShaperCaches == null)
                _threadLocalShaperCaches = new Dictionary<OpenTypeFontEngine, Dictionary<string, TextShaper>>();

            Dictionary<string, TextShaper> map;
            if (!_threadLocalShaperCaches.TryGetValue(this, out map))
            {
                map = new Dictionary<string, TextShaper>();
                _threadLocalShaperCaches[this] = map;
            }
            return map;
        }

        internal static List<string> GetLocationsCollection(
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories)
        {
            return DefaultFontLocations.GetLocationsCollection(fontDirectories, searchSystemDirectories);
        }

        // I OpenTypeFontEngine
        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException("OpenTypeFontEngine");
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            lock (_syncRoot)
            {
                _fontCache.Clear();
                _fontLocks.Clear();
            }

            // Best-effort cleanup of this engine's shaper map on the disposing thread.
            // Maps on other threads will be cleaned up when those threads next touch their map,
            // since they will see no entry for this engine and rebuild fresh. The lingering
            // entries are small (empty dictionaries) and held only by the disposing-time references.
            if (_threadLocalShaperCaches != null)
                _threadLocalShaperCaches.Remove(this);
        }
    }
}