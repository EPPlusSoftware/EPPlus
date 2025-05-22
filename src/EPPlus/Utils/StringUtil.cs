/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2025         EPPlus Software AB       EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class StringUtil
    {
        internal static string UnicodeSubstring(string input, int length)
        {
            if (string.IsNullOrEmpty(input) || length <= 0) return string.Empty;

            List<char> result = new List<char>();
            int count = 0;

            for (int i = 0; i < input.Length; i++)
            {
                if (count >= length) break; // Stop once we've extracted the required number of Unicode characters

                char c = input[i];

                // Handle surrogate pairs correctly
                if (char.IsHighSurrogate(c) && i + 1 < input.Length && char.IsLowSurrogate(input[i + 1]))
                {
                    if (count + 1 <= length) // Ensure the full Unicode character fits within the limit
                    {
                        result.Add(c);
                        result.Add(input[i + 1]);
                        count++;
                    }
                    break; // Stop after adding the surrogate pair if it exceeds the length
                }
                else
                {
                    // Single-character Unicode code point
                    result.Add(c);
                    count++;
                }
            }

            return new string(result.ToArray());
        }

        internal static string UnicodeSubstring(string input, int start, int length)
        {
            if (string.IsNullOrEmpty(input) || length <= 0 || start < 0) return string.Empty;

            List<char> result = new List<char>();
            int count = 0;
            int unicodeIndex = 0; // Tracks full Unicode character positions

            for (int i = 0; i < input.Length; i++)
            {
                if (unicodeIndex < start)
                {
                    // Skip characters until we reach the start position
                    char c = input[i];

                    if (char.IsHighSurrogate(c) && i + 1 < input.Length && char.IsLowSurrogate(input[i + 1]))
                        i++; // Move past surrogate pairs correctly

                    unicodeIndex++;
                    continue;
                }

                if (count >= length) break; // Stop once we extract enough Unicode characters

                char c2 = input[i];

                // Handle surrogate pairs properly
                if (char.IsHighSurrogate(c2) && i + 1 < input.Length && char.IsLowSurrogate(input[i + 1]))
                {
                    if (count + 1 <= length) // Ensure we don't exceed the requested length
                    {
                        result.Add(c2);
                        result.Add(input[i + 1]);
                        count++;
                    }
                    break; // Stop if we exceed the desired length
                    i++; // Skip next char (part of the surrogate pair)
                }
                else
                {
                    result.Add(c2);
                    count++;
                }
            }

            return new string(result.ToArray());
        }

        internal static int Compare(string l, string r, StringComparison comparison, ParsingContext context)
        {
            if(context.Configuration.EnableUnicodeAwareStringOperations)
            {
                return CompareStringUnicode(l, r);
            }
            return string.Compare(l, r, comparison);
        }

        internal static int CompareStringUnicode(string s1, string s2)
        {
            // Convert strings to full Unicode code points
            var s1Codes = EnumerateRunes(s1).ToArray();
            var s2Codes = EnumerateRunes(s2).ToArray();

            // Compare character by character based on their Unicode values
            int minLength = Math.Min(s1Codes.Length, s2Codes.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (s1Codes[i] != s2Codes[i])
                    return s1Codes[i].CompareTo(s2Codes[i]);
            }

            // If strings are identical up to minLength, compare by length
            return s1Codes.Length.CompareTo(s2Codes.Length);
        }

        static IEnumerable<int> EnumerateRunes(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];

                // Check if this is a high surrogate (start of a pair)
                if (char.IsHighSurrogate(c) && i + 1 < input.Length && char.IsLowSurrogate(input[i + 1]))
                {
                    // Combine high and low surrogates into a full Unicode code point
                    yield return char.ConvertToUtf32(c, input[i + 1]);
                    i++; // Skip the next character since it's part of the surrogate pair
                }
                else
                {
                    // Single character (not part of a surrogate pair)
                    yield return c;
                }
            }
        }
    }
}
