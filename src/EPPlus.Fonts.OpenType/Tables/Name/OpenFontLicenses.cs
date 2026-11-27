/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/19/2026         EPPlus Software AB           Open licence detection for web font embedding
 *************************************************************************************************/
using System;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    /// <summary>
    /// The font licences EPPlus accepts as open, so that a font under one of them may be embedded
    /// as a web font without the caller confirming it. Matched as case-insensitive substrings of the
    /// licence description (nameID 13) and the licence URL (nameID 14).
    ///
    /// Commercial fonts (Microsoft, Monotype, Adobe and fonts bought by customers) carry their own
    /// licence texts, which match none of these, and are never embedded by default.
    /// </summary>
    internal static class OpenFontLicenses
    {
        /// <summary>
        /// Phrases in nameID 13 that identify an open licence.
        /// </summary>
        internal static readonly string[] DescriptionPhrases =
        {
            "SIL Open Font License",
            "Apache License",
            "Ubuntu Font Licence",
            "Ubuntu Font License",
            "Bitstream Vera",
        };

        /// <summary>
        /// URL fragments in nameID 14 that identify an open licence.
        /// </summary>
        internal static readonly string[] UrlFragments =
        {
            "scripts.sil.org/OFL",
            "openfontlicense.org",
            "apache.org/licenses",
        };

        internal static bool IsOpenLicenseDescription(string description)
        {
            return ContainsAny(description, DescriptionPhrases);
        }

        internal static bool IsOpenLicenseUrl(string url)
        {
            return ContainsAny(url, UrlFragments);
        }

        private static bool ContainsAny(string value, string[] candidates)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            foreach (var candidate in candidates)
            {
                if (value.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
    }
}