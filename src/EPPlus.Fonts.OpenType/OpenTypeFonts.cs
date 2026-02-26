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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontCache;
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Scanner;
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
        private static volatile IFontResolver _fontResolver = new DefaultFontResolver();

        public static void Configure(IFontResolver resolver)
        {
            _fontResolver = resolver ?? throw new ArgumentNullException("resolver");
        }

        public static void Configure(FontResolutionConfig config)
        {
            _fontResolver = new DefaultFontResolver(config: config);
        }

        /// <summary>
        /// Clears all cached fonts and font locks.
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
        }

        /// <summary>
        /// Loads a font by name and subfamily, with thread-safe caching.
        /// When fontDirectories or searchSystemDirectories are specified, they take precedence
        /// over the globally configured resolver for this call only — thread-safe.
        /// </summary>
        /// <param name="fontName">Font family name</param>
        /// <param name="subFamily">Font subfamily (Regular, Bold, Italic, etc.)</param>
        /// <param name="fontDirectories">Additional directories to search. If null, uses globally configured resolver.</param>
        /// <param name="searchSystemDirectories">Whether to search system font directories</param>
        /// <param name="ignoreCache">If true, bypasses cache and loads font directly</param>
        /// <returns>Loaded OpenTypeFont or null if not found</returns>
        public static OpenTypeFont LoadFont(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true,
            bool ignoreCache = false)
        {
            // If caller specifies directories, use a local resolver for this call only.
            // This is thread-safe since the resolver is a local variable.
            var resolver = fontDirectories != null
                ? new DefaultFontResolver(fontDirectories, searchSystemDirectories)
                : _fontResolver;

            if (ignoreCache)
                return ResolveAndCreate(resolver, fontName, subFamily);

            // Cache key includes directories to avoid collisions between different resolver configs
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
        /// Builds a cache key that uniquely identifies a font + resolution config combination.
        /// </summary>
        internal static string BuildCacheKey(
            string fontName,
            FontSubFamily subFamily,
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories)
        {
            if (fontDirectories == null)
                return string.Format("{0}_{1}", fontName, subFamily);

            // Include directories in key to avoid collisions
            var dirs = string.Join("|", fontDirectories?.ToArray() ?? new string[] {"<empty>"});
            return string.Format("{0}_{1}_{2}_{3}", fontName, subFamily, dirs, searchSystemDirectories);
        }

        /// <summary>
        /// Resolves font bytes via the given resolver and creates an OpenTypeFont instance.
        /// Font format (TTF/OTF) is detected automatically from the SFNT header.
        /// </summary>
        private static OpenTypeFont ResolveAndCreate(IFontResolver resolver, string fontName, FontSubFamily subFamily)
        {
            var bytes = resolver.ResolveFont(fontName, subFamily);
            if (bytes == null)
                return null;

            return new OpenTypeFont(bytes);
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
                        var format = (ext == ".otf" || ext == ".cff") ? FontFormat.Otf : FontFormat.Ttf;
                        if (format != formatTarget.Value)
                            continue;
                    }
                }

                try
                {
                    var font = OpenTypeFontFactory.CreateFromFace(face);
                    if (font != null)
                        result.Add(font);
                    else
                        failures++;
                }
                catch (Exception ex) when (
                    ex is IOException ||
                    ex is UnauthorizedAccessException ||
                    ex is InvalidOperationException ||
                    ex is ArgumentException ||
                    ex is NotSupportedException ||
                    ex is EndOfStreamException)
                {
                    failures++;
                }
                catch (Exception ex)
                {
                    failures++;
                    System.Diagnostics.Debug.WriteLine(
                        string.Format(
                            "[OpenTypeFonts] UNEXPECTED ERROR loading font: {0} [TTC offset: {1}]\r\n" +
                            "  Exception: {2}\r\n" +
                            "  Message: {3}\r\n" +
                            "  Stack: {4}",
                            face.FilePath,
                            face.OffsetInFile,
                            ex.GetType().Name,
                            ex.Message,
                            ex.StackTrace));
                }
            }

            if (failures > 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    string.Format(
                        "[OpenTypeFonts] GetAllBaseFontData completed. Loaded {0} fonts, skipped {1} due to errors.",
                        result.Count,
                        failures));
            }

            return result;
        }

        /// <summary>
        /// Creates an OpenTypeFont from raw font bytes.
        /// Font format (TTF/OTF) is detected automatically from the SFNT header.
        /// </summary>
        public static OpenTypeFont GetFromBytes(byte[] bytes)
        {
            return new OpenTypeFont(bytes);
        }
    }
}