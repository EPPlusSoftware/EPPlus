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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Scanner;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    /// <summary>
    /// Default IFontResolver implementation that resolves fonts from the file system.
    /// Searches additional font directories and optionally system font directories.
    /// Supports fallback font chains via FontResolutionConfig.
    /// </summary>
    internal class DefaultFontResolver : IFontResolver
    {
        private readonly IEnumerable<string> _fontDirectories;
        private readonly bool _searchSystemDirectories;
        private readonly FontResolutionConfig _config;

        public DefaultFontResolver(
            IEnumerable<string> fontDirectories = null,
            bool searchSystemDirectories = true,
            FontResolutionConfig config = null)
        {
            _fontDirectories = fontDirectories ?? Enumerable.Empty<string>();
            _searchSystemDirectories = searchSystemDirectories;
            _config = config;
        }

        public byte[] ResolveFont(string fontName, FontSubFamily subFamily)
        {
            // Try exact match first
            var face = FontScannerV2.FindBestMatch(
                _fontDirectories, fontName, subFamily, _searchSystemDirectories);

            if (face != null && face.IsExactMatch)
                return ReadFontBytes(face.FilePath);

            // No exact match — try fallback chain
            if (_config != null)
            {
                var fallbacks = _config.GetFallbacks(fontName);
                if (fallbacks != null)
                {
                    foreach (var fallbackName in fallbacks)
                    {
                        var fallbackFace = FontScannerV2.FindBestMatch(
                            _fontDirectories, fallbackName, subFamily, _searchSystemDirectories);

                        if (fallbackFace != null && fallbackFace.IsExactMatch)
                            return ReadFontBytes(fallbackFace.FilePath);
                    }
                }
            }

            return null;
        }

        private static byte[] ReadFontBytes(string filePath)
        {
            return File.ReadAllBytes(filePath);
        }
    }
}