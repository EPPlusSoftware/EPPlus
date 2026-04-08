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
  01/10/2026         EPPlus Software AB           Fix threading issue with global lock
  01/23/2026         EPPlus Software AB           Improved thread-safety with per-font locking
  02/26/2026         EPPlus Software AB           Moved caching from DefaultFontResolver to here
  02/27/2026         EPPlus Software AB           Replaced Configure overloads with IEpplusFontConfiguration
  03/20/2026         EPPlus Software AB           Added thread-local TextShaper cache
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontCache;
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    public static class OpenTypeFonts
    {
        private static readonly object _syncRoot = new object();
        private static readonly Dictionary<string, object> _fontLocks = new Dictionary<string, object>();

        // Active resolver — replaced on Reset() and SetFontResolver().
        private static volatile IFontResolver _fontResolver;
        private static volatile bool _hasCustomResolver;

        // Singleton configuration instance. Internal events wired up in the static constructor.
        private static readonly EpplusFontConfiguration _configuration;

        // Thread-local TextShaper cache — each thread gets its own dictionary of shapers,
        // one per unique (fontName, subFamily) combination. The underlying OpenTypeFont instances
        // are shared via the global font cache, but TextShaper is not thread-safe to share.
        // NOTE: [ThreadStatic] field initializers only run on the primary thread. All other
        // threads will see null here — the null-check in GetTextShaper() handles this.
        [ThreadStatic]
        private static Dictionary<string, TextShaper> _threadLocalShaperCache;

        static OpenTypeFonts()
        {
            _configuration = new EpplusFontConfiguration();

            _configuration.OnReset += () =>
            {
                _hasCustomResolver = false;
                _fontResolver = new DefaultFontResolver();
                ClearFontCache();
            };

            _configuration.OnSetFontResolver += resolver =>
            {
                _hasCustomResolver = true;
                _fontResolver = resolver;
                ClearFontCache();
            };

            // Default resolver — no fallback config yet.
            _fontResolver = new DefaultFontResolver();
        }

        // -----------------------------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// The single entry point for configuring font behaviour.
        /// Changes are global and persist for the lifetime of the application unless
        /// <see cref="IEpplusFontConfiguration.Reset"/> is called inside the lambda.
        /// </summary>
        /// <example>
        /// // Simple fallback chain:
        /// OpenTypeFonts.Configure(config =>
        /// {
        ///     config.AddFallback("Arial", "Helvetica", "Roboto");
        /// });
        ///
        /// // Custom resolver (built-in Archivo Narrow fallback is bypassed):
        /// OpenTypeFonts.Configure(config =>
        /// {
        ///     config.SetFontResolver(new MyDatabaseFontResolver());
        /// });
        ///
        /// // Reset to factory defaults:
        /// OpenTypeFonts.Configure(config => config.Reset());
        /// </example>
        public static void Configure(Action<IEpplusFontConfiguration> configure)
        {
            if (configure == null)
                throw new ArgumentNullException("configure");

            configure(_configuration);

            // If the user did not install a custom resolver via SetFontResolver(), rebuild
            // DefaultFontResolver so it picks up any newly added fallback chains.
            if (!_hasCustomResolver)
                _fontResolver = new DefaultFontResolver(config: _configuration);
        }

        /// <summary>
        /// Gets a TextShaper for the given font, reusing a thread-local cached instance.
        /// The underlying OpenTypeFont is shared globally, but each thread gets its own
        /// TextShaper instance, ensuring thread safety without locking.
        /// Returns null if the font cannot be resolved.
        /// </summary>
        /// <param name="fontName">Font family name</param>
        /// <param name="subFamily">Font subfamily (Regular, Bold, Italic, etc.)</param>
        /// <param name="fontDirectories">Additional directories to search. If null, uses globally configured resolver.</param>
        /// <param name="searchSystemDirectories">Whether to search system font directories</param>
        public static TextShaper GetTextShaper(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true)
        {
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            // [ThreadStatic] fields are null on all threads except the primary — initialize on first use.
            if (_threadLocalShaperCache == null)
                _threadLocalShaperCache = new Dictionary<string, TextShaper>();

            // Use the same cache key logic as LoadFont to avoid collisions between
            // calls with and without explicit font directories.
            string key = BuildCacheKey(fontName, subFamily, fontDirectories, searchSystemDirectories);

            TextShaper shaper;
            if (!_threadLocalShaperCache.TryGetValue(key, out shaper))
            {
                var font = LoadFont(fontName, subFamily, fontDirectories, searchSystemDirectories);
                if (font == null)
                    return null;

                shaper = new TextShaper(font);
                _threadLocalShaperCache[key] = shaper;
            }

            return shaper;
        }

        public static TextLayoutEngine GetTextLayoutEngine(string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true)
        {
            var shaper = GetTextShaper(fontName, subFamily, fontDirectories, searchSystemDirectories);
            //TODO: Create layoutEngineCache in the style of shaperCache
            var layoutEngine = new TextLayoutEngine(shaper);

            return layoutEngine;
        }

        /// <summary>
        /// Clears all cached fonts, font locks and thread-local TextShaper cache.
        /// Thread-safe operation.
        /// </summary>
        public static void ClearFontCache()
        {
            lock (_syncRoot)
            {
                OpenTypeFontCache.Clear();
                FontScannerCache.Clear();
                _fontLocks.Clear();
            }

            // Clear the TextShaper cache for the calling thread. Other threads will
            // lazily rebuild their own caches on next call to GetTextShaper().
            _threadLocalShaperCache = null;
        }

        // -----------------------------------------------------------------------------------------
        // Internal font loading
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Loads a font by name and subfamily, with thread-safe caching.
        /// When fontDirectories or searchSystemDirectories are specified, they take precedence
        /// over the globally configured resolver for this call only — thread-safe.
        /// </summary>
        public static OpenTypeFont LoadFont(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true,
            bool ignoreCache = false)
        {
            var resolver = fontDirectories != null
                ? new DefaultFontResolver(fontDirectories, searchSystemDirectories)
                : _fontResolver;

            if (ignoreCache)
                return ResolveAndCreate(resolver, fontName, subFamily);

            string lockKey = BuildCacheKey(fontName, subFamily, fontDirectories, searchSystemDirectories);
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
                var cached = OpenTypeFontCache.GetFromCache(lockKey);
                if (cached != null && cached.Font != null && cached.IsLoaded)
                {
                    cached.Font.EnsureFullyLoaded();
                    return cached.Font;
                }

                OpenTypeFontCache.BeginCache(lockKey);

                var font = ResolveAndCreate(resolver, fontName, subFamily);
                if (font == null)
                    return null;

                font.EnsureFullyLoaded();
                font.IsReadOnly = true;
                OpenTypeFontCache.AddToCache(font, lockKey);
                return font;
            }
        }

        /// <summary>
        /// Returns all available font faces as fully loaded OpenTypeFont instances.
        /// Skips corrupt or unreadable fonts, but logs detailed information for diagnostics.
        /// This method is NOT cached and may take significant time to complete.
        /// </summary>
        public static List<OpenTypeFont> GetAllBaseFontData(
            List<string> fontDirectories,
            bool searchSystemDirectories = true,
            FontFormat? formatTarget = null)
        {
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
                        string.Format("[OpenTypeFonts] Failed to load font: {0} → {1}: {2}",
                            face.FilePath, ex.GetType().Name, ex.Message));
                }
            }

            if (failures > 0)
                System.Diagnostics.Debug.WriteLine(
                    string.Format("[OpenTypeFonts] {0} font(s) failed to load.", failures));

            return result;
        }

        /// <summary>
        /// Creates an OpenTypeFont directly from raw font bytes.
        /// Font format (TTF/OTF) is detected automatically from the SFNT header.
        /// </summary>
        /// <param name="bytes">Raw TTF/OTF font bytes.</param>
        /// <returns>A fully loaded OpenTypeFont instance.</returns>
        public static OpenTypeFont GetFromBytes(byte[] bytes)
        {
            if (bytes == null)
                throw new ArgumentNullException("bytes");

            var font = new OpenTypeFont(bytes);
            font.EnsureFullyLoaded();
            return font;
        }

        // -----------------------------------------------------------------------------------------
        // Internal helpers
        // -----------------------------------------------------------------------------------------

        internal static string BuildCacheKey(
            string fontName,
            FontSubFamily subFamily,
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories)
        {
            if (fontDirectories == null)
                return string.Format("{0}_{1}", fontName, subFamily);

            var dirs = string.Join("|", fontDirectories.ToArray());
            return string.Format("{0}_{1}_{2}_{3}", fontName, subFamily, dirs, searchSystemDirectories);
        }

        private static OpenTypeFont ResolveAndCreate(IFontResolver resolver, string fontName, FontSubFamily subFamily)
        {
            var bytes = resolver.ResolveFont(fontName, subFamily);
            if (bytes == null)
                return null;

            return new OpenTypeFont(bytes);
        }

        internal static List<string> GetLocationsCollection(
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories)
        {
            return DefaultFontLocations.GetLocationsCollection(fontDirectories, searchSystemDirectories);
        }
    }
}