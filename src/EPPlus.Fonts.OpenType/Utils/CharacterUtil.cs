/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/23/2025         EPPlus Software AB           ArrayPoolHelper implementation
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Utils
{
    internal class CharacterUtil
    {
        /// <summary>
        /// Extracts Unicode code points from a char array.
        /// Correctly handles surrogate pairs.
        /// </summary>
        public static IEnumerable<int> ExtractCodePointsFromChars(char[] chars)
        {
            var codePoints = new HashSet<int>();

            int i = 0;
            while (i < chars.Length)
            {
                if (i < chars.Length - 1 && char.IsHighSurrogate(chars[i]))
                {
                    char high = chars[i];
                    char low = chars[i + 1];

                    if (char.IsLowSurrogate(low))
                    {
                        // Valid surrogate pair - convert to code point
                        int codePoint = char.ConvertToUtf32(high, low);
                        codePoints.Add(codePoint);
                        i += 2; // Skip both surrogate chars
                        continue;
                    }
                    else
                    {
                        // Invalid - skip high surrogate alone
                        i++;
                        continue;
                    }
                }
                else if (char.IsLowSurrogate(chars[i]))
                {
                    // Lone low surrogate - skip it
                    i++;
                    continue;
                }
                else
                {
                    // Normal BMP character
                    codePoints.Add(chars[i]);
                    i++;
                }
            }

            return codePoints;
        }

        /// <summary>
        /// Extracts Unicode code points from a string.
        /// Correctly handles surrogate pairs.
        /// </summary>
        public static IEnumerable<int> ExtractCodePointsFromString(string text)
        {
            var codePoints = new HashSet<int>();

            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]))
                {
                    if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                    {
                        // Valid surrogate pair
                        int codePoint = char.ConvertToUtf32(text[i], text[i + 1]);
                        codePoints.Add(codePoint);
                        i++; // Skip low surrogate in next iteration
                    }
                    // else: Invalid high surrogate alone - skip it
                }
                else if (char.IsLowSurrogate(text[i]))
                {
                    // Lone low surrogate - skip it
                }
                else
                {
                    // Normal BMP character
                    codePoints.Add(text[i]);
                }
            }

            return codePoints;
        }
    }
}
