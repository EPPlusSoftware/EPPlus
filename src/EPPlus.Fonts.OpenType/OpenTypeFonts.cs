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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontCache;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Utils.Platform;
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

        #region --- Platform-specific font locations (unchanged, beautiful as always) ---

        private static string GetWindowsFolder()
        {
            var winfolder = @"c:\Windows";
#if !NET35
            var wf = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(wf) && Directory.Exists(wf)) return wf;
#endif
            var ewf = Environment.GetEnvironmentVariable("WINDIR");
            if (!string.IsNullOrEmpty(ewf) && Directory.Exists(ewf))
            {
                winfolder = ewf;
            }
            return winfolder;
        }

        internal static readonly List<string> winFontLocations = new List<string>()
        {
            Path.Combine(GetWindowsFolder(), "Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\FontCache"),
        };

        internal static readonly List<string> linFontLocations = new List<string>()
        {
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            "/usr/share/X11/fonts",
            Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".fonts"),
            Path.Combine(Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".local"), "share"), "fonts"),
        };

        internal static readonly List<string> macFontLocations = new List<string>()
        {
            "/System/Library/Fonts",
            "/Library/Fonts",
            "/Network/Library/Fonts",
            Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", "Library"), "Fonts"),
        };

        internal static List<string> GetLocationsCollection(IEnumerable<string> fontDirectories, bool searchSystemDirectories = true)
        {
            var fontLocations = new List<string>();
            fontLocations.AddRange(fontDirectories ?? Enumerable.Empty<string>());

            if (searchSystemDirectories)
            {
                var platform = PlatformUtils.GetPlatform();
                if (platform == PlatformUtils.OperatingSystem.Windows)
                    fontLocations.AddRange(winFontLocations);
                else if (platform == PlatformUtils.OperatingSystem.Mac)
                    fontLocations.AddRange(macFontLocations);
                else
                    fontLocations.AddRange(linFontLocations);
            }

            return fontLocations;
        }

        #endregion

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
        /// Returns a fully loaded OpenTypeFont – fast, cached, safe.
        /// Uses FontScannerV2 under the hood.
        /// </summary>
        public static OpenTypeFont GetFontDataOpen(
            IEnumerable<string> fontDirectories,
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            bool searchSystemDirectories = true,
            bool ignoreCache = false)
        {
            // Create per-font lock key
            string lockKey = $"{fontName}_{subFamily}";
            object fontLock;

            lock (_syncRoot)
            {
                if (!_fontLocks.TryGetValue(lockKey, out fontLock))
                {
                    fontLock = new object();
                    _fontLocks[lockKey] = fontLock;
                }
            }

            // Now lock PER FONT, not globally
            lock (fontLock)
            {
                if (!ignoreCache)
                {
                    if (OpenTypeFontCache.Contains(fontName, subFamily))
                    {
                        var cached = OpenTypeFontCache.GetFromCache(fontName, subFamily);
                        if (cached?.Font != null && cached.IsLoaded)
                            return cached.Font;
                    }
                    OpenTypeFontCache.BeginCache(fontName, subFamily);
                }

                var face = FontScannerV2.FindBestMatch(fontDirectories, fontName, subFamily, searchSystemDirectories);
                if (face == null)
                    return null;

                var font = OpenTypeFontFactory.CreateFromFace(face);

                if (!ignoreCache)
                    OpenTypeFontCache.AddToCache(font, fontName, subFamily);

                return font;
            }
        }

        /// <summary>
        /// Legacy wrapper – kept for backward compatibility.
        /// </summary>
        public static OpenTypeFont GetFontData(
            IEnumerable<string> fontDirectories,
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            bool searchSystemDirectories = true,
            bool ignoreCache = false)
        {
            return GetFontDataOpen(fontDirectories, fontName, subFamily, searchSystemDirectories, ignoreCache);
        }

        /// <summary>
        /// Returns all available font faces as fully loaded OpenTypeFont instances.
        /// Skips corrupt or unreadable fonts, but logs detailed information for diagnostics.
        /// </summary>
        public static List<OpenTypeFont> GetAllBaseFontData(
            List<string> fontDirectories,
            bool searchSystemDirectories = true,
            FontFormat? formatTarget = null)
        {
            var locations = GetLocationsCollection(fontDirectories, searchSystemDirectories);
            var faces = FontScannerV2.EnumerateAllFaces(locations);

            var result = new List<OpenTypeFont>(faces.Count);
            var failures = 0;

            foreach (var face in faces)
            {
                // Filter by format if requested
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
                    OpenTypeFont font = OpenTypeFontFactory.CreateFromFace(face);
                    if (font != null)
                    {
                        result.Add(font);
                    }
                    else
                    {
                        failures++;
                    }
                }
                catch (Exception ex) when (
                    ex is IOException ||
                    ex is UnauthorizedAccessException ||
                    ex is InvalidOperationException ||
                    ex is ArgumentException ||
                    ex is NotSupportedException ||
                    ex is EndOfStreamException)
                {
                    // These are expected for corrupt or inaccessible fonts
                    failures++;
                }
                catch (Exception ex)
                {
                    // Unexpected exceptions – log with full details (never swallow these silently)
                    failures++;
                    System.Diagnostics.Debug.WriteLine(
                        $"[OpenTypeFonts] UNEXPECTED ERROR loading font: {face.FilePath} [TTC offset: {face.OffsetInFile}]\r\n" +
                        $"  Exception: {ex.GetType().Name}\r\n" +
                        $"  Message: {ex.Message}\r\n" +
                        $"  Stack: {ex.StackTrace}");
                }
            }

            if (failures > 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[OpenTypeFonts] GetAllBaseFontData completed. Loaded {result.Count} fonts, skipped {failures} due to errors.");
            }

            return result;
        }
    }
}