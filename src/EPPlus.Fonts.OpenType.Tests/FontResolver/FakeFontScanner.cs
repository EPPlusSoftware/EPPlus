/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/06/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Test fake of IFontScanner. Returns predefined FontFaceInfo objects for registered
    /// (familyName, subFamily) pairs. Lookup is case-insensitive on family name. Unregistered
    /// requests return null.
    /// </summary>
    internal sealed class FakeFontScanner : IFontScanner
    {
        private readonly Dictionary<string, FontFaceInfo> _registry =
            new Dictionary<string, FontFaceInfo>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers that a request for the given family + subfamily should resolve to the
        /// given file path. Sets IsExactMatch = true on the returned face. Returns this for
        /// fluent chaining.
        /// </summary>
        public FakeFontScanner Register(string familyName, FontSubFamily subFamily, string filePath)
        {
            _registry[MakeKey(familyName, subFamily)] = new FontFaceInfo
            {
                FilePath = filePath,
                FamilyName = familyName,
                Subfamily = subFamily,
                IsExactMatch = true,
            };
            return this;
        }

        public FontFaceInfo? FindBestMatch(
            IEnumerable<string> additionalDirectories,
            string familyName,
            FontSubFamily desiredStyle,
            bool searchSystemDirectories)
        {
            if (string.IsNullOrEmpty(familyName))
                return null;

            return _registry.TryGetValue(MakeKey(familyName, desiredStyle), out var face)
                ? face
                : null;
        }

        private static string MakeKey(string familyName, FontSubFamily subFamily)
        {
            return familyName + "|" + subFamily;
        }
    }
}