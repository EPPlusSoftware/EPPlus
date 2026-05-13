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
  05/06/2026         EPPlus Software AB           Transactional Configure; single resolver, single cache key
  05/13/2026         EPPlus Software AB           Reduced to a thin facade over a singleton OpenTypeFontEngine
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Static facade for the OpenType font system. Delegates to a singleton
    /// <see cref="OpenTypeFontEngine"/> instance for backward compatibility with
    /// callers that have not yet been migrated to per-instance engine usage.
    ///
    /// New code should prefer creating and owning an OpenTypeFontEngine directly.
    /// </summary>
    public static class OpenTypeFonts
    {
        private static readonly object _syncRoot = new object();
        private static OpenTypeFontEngine _default = new OpenTypeFontEngine();

        // -----------------------------------------------------------------------------------------
        // Configuration (mutates the singleton engine)
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Reconfigures the singleton font engine used by this static facade.
        /// Internally creates a new <see cref="OpenTypeFontEngine"/> with the supplied
        /// configuration and replaces the previous singleton. The previous engine is disposed,
        /// invalidating any caches built against it.
        ///
        /// This method exists for source compatibility with callers that have not yet been
        /// migrated to per-instance engine usage. New code should create its own
        /// <see cref="OpenTypeFontEngine"/> instead.
        /// </summary>
        public static void Configure(Action<IEpplusFontConfiguration> configure)
        {
            if (configure == null)
                throw new ArgumentNullException("configure");

            OpenTypeFontEngine oldEngine;
            lock (_syncRoot)
            {
                oldEngine = _default;
                _default = new OpenTypeFontEngine(configure);
            }

            // Dispose the old engine outside the lock. Any caller still holding a reference
            // to it (e.g. via GetTextShaper) will get an ObjectDisposedException on next use,
            // which is the intended signal that configuration changed underneath them.
            try { oldEngine.Dispose(); } catch { /* swallow - best effort */ }
        }

        // -----------------------------------------------------------------------------------------
        // Delegating API
        // -----------------------------------------------------------------------------------------

        public static TextShaper GetTextShaper(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            return _default.GetTextShaper(fontName, subFamily);
        }

        public static TextLayoutEngine GetTextLayoutEngine(string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            return _default.GetTextLayoutEngine(fontName, subFamily);
        }

        public static TextLayoutEngine GetTextLayoutEngineForFont(MeasurementFont font)
        {
            return _default.GetTextLayoutEngineForFont(font);
        }

        public static ITextShaper GetShaperForFont(MeasurementFont font)
        {
            return _default.GetShaperForFont(font);
        }

        public static FontSubFamily GetFontSubFamily(MeasurementFontStyles style)
        {
            return OpenTypeFontEngine.GetFontSubFamily(style);
        }

        public static void ClearFontCache()
        {
            _default.ClearFontCache();
        }

        public static OpenTypeFont LoadFont(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular,
            bool ignoreCache = false)
        {
            return _default.LoadFont(fontName, subFamily, ignoreCache);
        }

        /// <summary>
        /// Overload preserved for source compatibility with callers that pass fontDirectories.
        /// The directories argument is IGNORED - to add directories permanently, configure the
        /// engine via <see cref="Configure"/> or create your own <see cref="OpenTypeFontEngine"/>.
        /// </summary>
        [Obsolete("Pass font directories through OpenTypeFonts.Configure(cfg => cfg.FontDirectories.Add(...)) or use OpenTypeFontEngine directly. The directories argument is ignored.", false)]
        public static OpenTypeFont LoadFont(
            string fontName,
            FontSubFamily subFamily,
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories = true,
            bool ignoreCache = false)
        {
            return _default.LoadFont(fontName, subFamily, ignoreCache);
        }

        public static List<OpenTypeFont> GetAllBaseFontData(
            List<string> fontDirectories,
            bool searchSystemDirectories = true,
            FontFormat? formatTarget = null)
        {
            return _default.GetAllBaseFontData(fontDirectories, searchSystemDirectories, formatTarget);
        }

        public static OpenTypeFont GetFromBytes(byte[] bytes)
        {
            return _default.GetFromBytes(bytes);
        }

        public static FontAvailability GetFontAvailability(
            string fontName,
            FontSubFamily subFamily = FontSubFamily.Regular)
        {
            return _default.GetFontAvailability(fontName, subFamily);
        }

        internal static string BuildCacheKey(string fontName, FontSubFamily subFamily)
        {
            return OpenTypeFontEngine.BuildCacheKey(fontName, subFamily);
        }

        internal static List<string> GetLocationsCollection(
            IEnumerable<string> fontDirectories,
            bool searchSystemDirectories)
        {
            return OpenTypeFontEngine.GetLocationsCollection(fontDirectories, searchSystemDirectories);
        }
    }
}