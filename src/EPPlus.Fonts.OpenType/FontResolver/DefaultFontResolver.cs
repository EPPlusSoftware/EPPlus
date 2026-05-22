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
 *************************************************************************************************/
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
    /// </summary>
    internal class DefaultFontResolver : IFontResolver, IFontAvailabilityProvider
    {
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
            // 1. Try exact match first
            var face = _scanner.FindBestMatch(
                _fontDirectories, fontName, subFamily, _searchSystemDirectories);

            if (face != null && face.IsExactMatch)
                return _fileReader.ReadFontBytes(face);

            // 2. No exact match — try user-configured fallback chain
            if (_config != null)
            {
                var userFallbacks = _config.GetFallbacks(fontName);
                if (userFallbacks != null)
                {
                    var resolved = TryResolveFromChain(userFallbacks, subFamily);
                    if (resolved != null)
                        return resolved;
                }
            }

            // 3. Try built-in fallback chain for known Office/system fonts.
            // Runs after user config so user preferences win, but still provides a metric-aware
            // safety net for fonts the user hasn't configured.
            var builtinFallbacks = BuiltinFontFallbackChains.GetFallbacks(fontName);
            if (builtinFallbacks != null)
            {
                var resolved = TryResolveFromChain(builtinFallbacks, subFamily);
                if (resolved != null)
                    return resolved;
            }

            // 4. No match found — fall back to built-in Archivo Narrow.
            // Only applies when using DefaultFontResolver (i.e. no custom resolver installed).
            return EmbeddedFonts.LoadArchivoNarrow(subFamily).RawData;
        }

        /// <summary>
        /// Attempts to resolve a font by walking through a chain of fallback names.
        /// Returns the bytes of the first chain entry that produces an exact match, or null if
        /// no entry resolves. Each entry is required to match the requested subFamily — falling
        /// back from a Bold request to a Regular face would defeat the purpose of fallback.
        /// </summary>
        private byte[] TryResolveFromChain(IEnumerable<string> chain, FontSubFamily subFamily)
        {
            foreach (var fallbackName in chain)
            {
                var fallbackFace = _scanner.FindBestMatch(
                    _fontDirectories, fallbackName, subFamily, _searchSystemDirectories);

                if (fallbackFace != null && fallbackFace.IsExactMatch)
                    return _fileReader.ReadFontBytes(fallbackFace);
            }
            return null;
        }
    }
}