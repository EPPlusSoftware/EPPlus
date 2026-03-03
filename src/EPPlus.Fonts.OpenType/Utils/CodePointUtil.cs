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
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Utils
{
    internal static class CodePointUtil
    {
        /// <summary>
        /// Extracts distinct Unicode code points from a sequence of chars,
        /// correctly combining surrogate pairs into supplementary plane code points.
        /// Lone surrogates are skipped.
        /// </summary>
        internal static IEnumerable<int> ExtractCodePoints(IEnumerable<char> chars)
        {
            var codePoints = new HashSet<int>();
            char? pendingHighSurrogate = null;

            foreach (var c in chars)
            {
                if (pendingHighSurrogate.HasValue)
                {
                    if (char.IsLowSurrogate(c))
                    {
                        // Valid pair - combine into code point
                        int cp = char.ConvertToUtf32(pendingHighSurrogate.Value, c);
                        codePoints.Add(cp);
                        pendingHighSurrogate = null;
                        continue;
                    }

                    // Previous high surrogate had no matching low - skip it
                    pendingHighSurrogate = null;
                }

                if (char.IsHighSurrogate(c))
                {
                    pendingHighSurrogate = c;
                }
                else if (!char.IsLowSurrogate(c))
                {
                    // Normal BMP character
                    codePoints.Add(c);
                }
                // Lone low surrogates are skipped
            }

            return codePoints;
        }

        /// <summary>
        /// Converts a set of Unicode code points to a string,
        /// correctly encoding supplementary plane characters as surrogate pairs.
        /// </summary>
        internal static string CodePointsToString(HashSet<int> codePoints)
        {
            var sb = new System.Text.StringBuilder(codePoints.Count * 2);

            foreach (var cp in codePoints)
            {
                if (cp <= 0xFFFF)
                {
                    sb.Append((char)cp);
                }
                else
                {
                    // Supplementary plane - encode as surrogate pair
                    sb.Append(char.ConvertFromUtf32(cp));
                }
            }

            return sb.ToString();
        }
    }
}
