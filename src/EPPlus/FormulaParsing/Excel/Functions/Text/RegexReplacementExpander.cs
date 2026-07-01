/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/6/2026         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    /// <summary>
    /// Expands the replacement string of REGEXREPLACE using Excel's replacement grammar, which
    /// differs from the grammar used by .NET's <see cref="Regex.Replace(string, string)"/>.
    ///
    /// The grammar was verified against Excel desktop:
    ///   Group references (greedy, consume all following digits):
    ///     \N  -> group N        (N must be 1..groupCount, \0 is invalid)
    ///     $N  -> group N        (N must be 0..groupCount, $0 is the whole match)
    ///   Literals:
    ///     $$              -> literal $
    ///     \\              -> literal \
    ///     \ + non-letter  -> the literal character (e.g. \, \- \. \$ \) \&lt;space&gt;)
    ///   Character escapes (-> the corresponding character):
    ///     \n \t \r \f \a \v \e \b, \cX (control), \xNN (hex)
    ///     \u is NOT supported and is silently dropped (Excel quirk: \u0041 -> "0041")
    ///   Invalid (-> #VALUE!):
    ///     \w \W \s \S \d \D and any other letter escape (cannot become a single character)
    ///     $ followed by anything other than a digit or $ (including a trailing $)
    ///   Edge:
    ///     a trailing lone \ is silently dropped
    /// </summary>
    internal static class RegexReplacementExpander
    {
        /// <summary>
        /// Validates a replacement template against Excel's grammar without needing a match.
        /// Returns false when the template should produce a #VALUE! error.
        /// </summary>
        /// <param name="template">The replacement string.</param>
        /// <param name="maxGroup">The highest capturing group number in the pattern (0 if none).</param>
        public static bool IsValidTemplate(string template, int maxGroup)
        {
            return ScanAndExpand(template, maxGroup, null, null);
        }

        /// <summary>
        /// Expands a replacement template for a specific match. The template is assumed to have
        /// already passed <see cref="IsValidTemplate"/>.
        /// </summary>
        public static string Expand(Match match, string template, int maxGroup)
        {
            var output = new StringBuilder();
            ScanAndExpand(template, maxGroup, match, output);
            return output.ToString();
        }

        /// <summary>
        /// Single left-to-right scan that both validates and (when <paramref name="match"/> and
        /// <paramref name="output"/> are supplied) expands the template. Returns false on the
        /// first construct that Excel rejects with #VALUE!.
        /// </summary>
        private static bool ScanAndExpand(string template, int maxGroup, Match match, StringBuilder output)
        {
            if (template == null)
            {
                return true;
            }

            int i = 0;
            int n = template.Length;
            while (i < n)
            {
                char ch = template[i];
                if (ch == '\\')
                {
                    if (i + 1 >= n)
                    {
                        // Trailing lone backslash -> silently dropped.
                        i++;
                        continue;
                    }

                    char next = template[i + 1];
                    if (next >= '0' && next <= '9')
                    {
                        // Greedy group reference \N. \0 is invalid, N must be 1..maxGroup.
                        int start = i + 1;
                        int j = start;
                        while (j < n && template[j] >= '0' && template[j] <= '9')
                        {
                            j++;
                        }
                        int groupNo = ParseGroupNumber(template, start, j);
                        if (groupNo < 1 || groupNo > maxGroup)
                        {
                            return false;
                        }
                        if (output != null)
                        {
                            output.Append(match.Groups[groupNo].Value);
                        }
                        i = j;
                    }
                    else if (next == '\\')
                    {
                        if (output != null)
                        {
                            output.Append('\\');
                        }
                        i += 2;
                    }
                    else if (IsControlEscapeLetter(next))
                    {
                        if (output != null)
                        {
                            output.Append(ControlEscapeChar(next));
                        }
                        i += 2;
                    }
                    else if (next == 'x')
                    {
                        // \xNN hex escape, up to two hex digits.
                        int start = i + 2;
                        int j = start;
                        while (j < n && j < start + 2 && IsHexDigit(template[j]))
                        {
                            j++;
                        }
                        if (j == start)
                        {
                            // \x with no hex digit.
                            return false;
                        }
                        if (output != null)
                        {
                            output.Append((char)HexValue(template, start, j));
                        }
                        i = j;
                    }
                    else if (next == 'c')
                    {
                        // \cX control escape; X must be a letter.
                        if (i + 2 >= n || !IsAsciiLetter(template[i + 2]))
                        {
                            return false;
                        }
                        if (output != null)
                        {
                            output.Append((char)(template[i + 2] & 0x1F));
                        }
                        i += 3;
                    }
                    else if (next == 'u')
                    {
                        // Excel does not support \u; it is silently dropped and the following
                        // characters remain literal (verified: \u0041 -> "0041").
                        i += 2;
                    }
                    else if (IsAsciiLetter(next))
                    {
                        // Any other letter escape (\w \W \s \S \d \D, ...) cannot become a single
                        // character -> #VALUE!.
                        return false;
                    }
                    else
                    {
                        // Backslash before a non-letter, non-digit -> the literal character.
                        if (output != null)
                        {
                            output.Append(next);
                        }
                        i += 2;
                    }
                }
                else if (ch == '$')
                {
                    if (i + 1 >= n)
                    {
                        // Lone trailing $ -> #VALUE!.
                        return false;
                    }

                    char next = template[i + 1];
                    if (next == '$')
                    {
                        if (output != null)
                        {
                            output.Append('$');
                        }
                        i += 2;
                    }
                    else if (next >= '0' && next <= '9')
                    {
                        // Greedy group reference $N. $0 is the whole match, N must be 0..maxGroup.
                        int start = i + 1;
                        int j = start;
                        while (j < n && template[j] >= '0' && template[j] <= '9')
                        {
                            j++;
                        }
                        int groupNo = ParseGroupNumber(template, start, j);
                        if (groupNo < 0 || groupNo > maxGroup)
                        {
                            return false;
                        }
                        if (output != null)
                        {
                            output.Append(match.Groups[groupNo].Value);
                        }
                        i = j;
                    }
                    else
                    {
                        // $ followed by anything other than a digit or $ -> #VALUE!.
                        return false;
                    }
                }
                else
                {
                    if (output != null)
                    {
                        output.Append(ch);
                    }
                    i++;
                }
            }

            return true;
        }

        private static int ParseGroupNumber(string s, int start, int end)
        {
            long acc = 0;
            for (int k = start; k < end; k++)
            {
                acc = acc * 10 + (s[k] - '0');
                if (acc > int.MaxValue)
                {
                    // Larger than any possible group number; the caller's range check rejects it.
                    return int.MaxValue;
                }
            }
            return (int)acc;
        }

        private static bool IsControlEscapeLetter(char c)
        {
            return c == 'n' || c == 't' || c == 'r' || c == 'f'
                || c == 'a' || c == 'v' || c == 'b' || c == 'e';
        }

        private static char ControlEscapeChar(char c)
        {
            switch (c)
            {
                case 'n': return '\n';
                case 't': return '\t';
                case 'r': return '\r';
                case 'f': return '\f';
                case 'a': return '\a';
                case 'v': return '\v';
                case 'b': return '\b';
                case 'e': return (char)27;
                default: return c;
            }
        }

        private static bool IsAsciiLetter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        private static bool IsHexDigit(char c)
        {
            return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
        }

        private static int HexValue(string s, int start, int end)
        {
            int value = 0;
            for (int k = start; k < end; k++)
            {
                value = value * 16 + HexDigit(s[k]);
            }
            return value;
        }

        private static int HexDigit(char c)
        {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            return c - 'A' + 10;
        }
    }
}