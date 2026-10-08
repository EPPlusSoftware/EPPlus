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
  02/26/2026         EPPlus Software AB           Removed caching (moved to OpenTypeFonts)
  02/27/2026         EPPlus Software AB           Replaced FontResolutionConfig with EpplusFontConfiguration, added Archivo Narrow fallback
  03/02/2026         EPPlus Software AB           TTC support: extract individual font from collection
  05/06/2026         EPPlus Software AB           Built-in fallback chains for common Office fonts
  05/06/2026         EPPlus Software AB           Extracted IFontScanner and IFontFileReader for testability
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Logging;
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    /// <summary>
    /// Default IFontResolver implementation that resolves fonts from the file system.
    /// Searches additional font directories and optionally system font directories.
    /// Supports fallback font chains via EpplusFontConfiguration as well as a built-in
    /// metric-aware fallback chain for common Office and system fonts.
    /// TTC (TrueType Collection) files are handled transparently by the IFontFileReader.
    ///
    /// Each resolution reports which step decided the outcome to the logger configured on
    /// <see cref="EpplusFontConfiguration"/>, if any.
    /// </summary>
    internal class DefaultFontResolver : IFontResolver, IFontAvailabilityProvider
    {
        private const string LastResortFamily = "Archivo Narrow";

        private readonly IEnumerable<string> _fontDirectories;
        private readonly bool _searchSystemDirectories;
        private readonly EpplusFontConfiguration _config;
        private readonly IFontScanner _scanner;
        private readonly IFontFileReader _fileReader;

        public DefaultFontResolver(
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true,
            EpplusFontConfiguration config = null,
            IFontScanner scanner = null,
            IFontFileReader fileReader = null)
        {
            _fontDirectories = fontDirectories ?? Enumerable.Empty<string>();
            _searchSystemDirectories = searchSystemDirectories;
            _config = config;
            _scanner = scanner ?? new DefaultFontScanner();
            _fileReader = fileReader ?? new DefaultFontFileReader();
        }

        public FontAvailability GetFontAvailability(string fontName, FontSubFamily subFamily)
        {
            if (string.IsNullOrEmpty(fontName))
                return FontAvailability.NotFound;

            // special case for Archivo Narrow which is distributed as last-resort-font with EPPlus
            if (string.Equals("archivo narrow", fontName, StringComparison.OrdinalIgnoreCase))
            {
                return FontAvailability.Exact;
            }

            var face = _scanner.FindBestMatch(
                _fontDirectories,
                fontName,
                subFamily,
                _searchSystemDirectories);

            if (face == null)
                return FontAvailability.NotFound;

            // FindBestMatch may return a non-matching face when no real match exists.
            // Verify the returned face actually belongs to the requested family.
            if (!string.Equals(face.FamilyName, fontName, StringComparison.OrdinalIgnoreCase))
                return FontAvailability.NotFound;

            return face.IsExactMatch
                ? FontAvailability.Exact
                : FontAvailability.FamilyOnly;
        }

        public byte[] ResolveFont(string fontName, FontSubFamily subFamily)
        {
            var logger = _config != null ? _config.ActiveLogger : NullFontLogger.Instance;

            // 1.  special case for Archivo Narrow which is distributed as last-resort-font with EPPlus
            if (string.Equals("archivo narrow", fontName, StringComparison.OrdinalIgnoreCase))
            {
                LogResolved(logger, fontName, subFamily, "the embedded last-resort font was requested");
                return EmbeddedFonts.LoadArchivoNarrow(subFamily).RawData;
            }

            // 2. Try exact match first
            var face = _scanner.FindBestMatch(
                _fontDirectories, fontName, subFamily, _searchSystemDirectories);

            if (face != null && face.IsExactMatch)
            {
                LogResolved(logger, fontName, subFamily, "exact match");
                return _fileReader.ReadFontBytes(face);
            }

            // 3. No exact match — try user-configured fallback chain
            string[] userFallbacks = null;
            if (_config != null)
            {
                userFallbacks = _config.GetFallbacks(fontName);
                if (userFallbacks != null)
                {
                    string matched;
                    var resolved = TryResolveFromChain(userFallbacks, subFamily, out matched);
                    if (resolved != null)
                    {
                        LogFallback(logger, fontName, subFamily, matched, "user-configured chain", userFallbacks);
                        return resolved;
                    }
                }
            }

            // 4. Try built-in fallback chain for known Office/system fonts.
            // Runs after user config so user preferences win, but still provides a metric-aware
            // safety net for fonts the user hasn't configured.
            var builtinFallbacks = BuiltinFontFallbackChains.GetFallbacks(fontName);
            if (builtinFallbacks != null)
            {
                string matched;
                var resolved = TryResolveFromChain(builtinFallbacks, subFamily, out matched);
                if (resolved != null)
                {
                    LogFallback(logger, fontName, subFamily, matched, "built-in chain", builtinFallbacks);
                    return resolved;
                }
            }

            // 5. No match found — fall back to built-in Archivo Narrow.
            // Only applies when using DefaultFontResolver (i.e. no custom resolver installed).
            LogLastResort(logger, fontName, subFamily, face != null ? face.FamilyName : null, userFallbacks, builtinFallbacks);
            return EmbeddedFonts.LoadArchivoNarrow(subFamily).RawData;
        }

        /// <summary>
        /// Attempts to resolve a font by walking through a chain of fallback names.
        /// Returns the bytes of the first chain entry that produces an exact match, or null if
        /// no entry resolves. Each entry is required to match the requested subFamily — falling
        /// back from a Bold request to a Regular face would defeat the purpose of fallback.
        /// <paramref name="matchedName"/> receives the chain entry that matched, or null.
        /// </summary>
        private byte[] TryResolveFromChain(IEnumerable<string> chain, FontSubFamily subFamily, out string matchedName)
        {
            foreach (var fallbackName in chain)
            {
                var fallbackFace = _scanner.FindBestMatch(
                    _fontDirectories, fallbackName, subFamily, _searchSystemDirectories);

                if (fallbackFace != null && fallbackFace.IsExactMatch)
                {
                    matchedName = fallbackName;
                    return _fileReader.ReadFontBytes(fallbackFace);
                }
            }

            matchedName = null;
            return null;
        }

        // -----------------------------------------------------------------------------------------
        // Logging
        // -----------------------------------------------------------------------------------------

        private static void LogResolved(IFontLogger logger, string fontName, FontSubFamily subFamily, string reason)
        {
            if (!FontLog.IsEnabled(logger, FontLogSeverity.Debug))
                return;

            FontLog.Write(
                logger,
                FontLogSeverity.Debug,
                FontLogEventType.FontResolved,
                string.Format("Font '{0}' {1}: {2}.", fontName, subFamily, reason),
                fontName,
                fontName);
        }

        private static void LogFallback(
            IFontLogger logger,
            string fontName,
            FontSubFamily subFamily,
            string matchedName,
            string chainKind,
            string[] chain)
        {
            if (!FontLog.IsEnabled(logger, FontLogSeverity.Information))
                return;

            FontLog.Write(
                logger,
                FontLogSeverity.Information,
                FontLogEventType.FontFallback,
                string.Format(
                    "Font '{0}' {1} -> '{2}' ({3}: {4}).",
                    fontName, subFamily, matchedName, chainKind, FontLog.JoinNames(chain)),
                fontName,
                matchedName);
        }

        private static void LogLastResort(
            IFontLogger logger,
            string fontName,
            FontSubFamily subFamily,
            string closestFamily,
            string[] userFallbacks,
            string[] builtinFallbacks)
        {
            if (!FontLog.IsEnabled(logger, FontLogSeverity.Warning))
                return;

            var message = string.Format(
                "Font '{0}' {1} -> '{2}' (last resort: no exact match; user chain: {3}; built-in chain: {4}).",
                fontName, subFamily, LastResortFamily,
                FontLog.JoinNames(userFallbacks), FontLog.JoinNames(builtinFallbacks));

            // The scanner can return a face that is not an exact match, for instance the right
            // family in another style. It is rejected on purpose, which is worth knowing.
            if (closestFamily != null)
            {
                message += string.Format(
                    " Closest face found, '{0}', was rejected because it is not an exact match.",
                    closestFamily);
            }

            FontLog.Write(
                logger,
                FontLogSeverity.Warning,
                FontLogEventType.FontLastResort,
                message,
                fontName,
                LastResortFamily);
        }
    }
}