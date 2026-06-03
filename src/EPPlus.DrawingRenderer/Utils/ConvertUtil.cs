/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;
namespace EPPlus.DrawingRenderer.Utils
{
    internal static class ConvertUtil
    {
        static class ParseArguments
        {
            public static NumberStyles Number = NumberStyles.Float | NumberStyles.AllowThousands;
            public static DateTimeStyles DateTime = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
        }
        internal static bool IsNumericOrDate(object candidate)
        {
            if (candidate == null) return false;
            var t = candidate.GetType();
            if (t.IsPrimitive) return true;
            return t == typeof(decimal) || t == typeof(DateTime) || t == typeof(TimeSpan);
        }
        internal static bool IsNumeric(object candidate)
        {
            if (candidate == null) return false;
            var t = candidate.GetType();
            if (t.IsPrimitive) return true;
            return t == typeof(decimal);
        }
        internal static bool IsExcelNumeric(object candidate)
        {
            if (candidate == null) return false;
            var t = candidate.GetType();
            if (t.IsPrimitive && t != typeof(bool)) return true;
            return t == typeof(decimal) || t == typeof(DateTime) || t == typeof(TimeSpan);
        }

        internal static bool IsNumericOrDate(object candidate, bool includeNumericString, bool includePercentageString)
        {
            if (IsNumericOrDate(candidate)) return true;
            if (candidate != null)
            {
                if (includeNumericString && TryParseNumericString(candidate.ToString(), out double d, CultureInfo.CurrentCulture))
                {
                    return true;
                }
                else if (includePercentageString && IsPercentageString(candidate.ToString()))
                {
                    return true;
                }
            }
            return false;
        }

        internal static bool IsPercentageString(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            s = s.Trim();
            if (s.Trim().EndsWith("%"))
            {
                var tmp = s;
                int n = 0;
                while (tmp.Length > 0 && tmp.Last() == '%')
                {
                    tmp = tmp.Substring(0, tmp.Length - 1);
                    n++;
                    if (n > 1)
                    {
                        return false;
                    }
                }
                return double.TryParse(tmp, out double n2);
            }
            return false;
        }
        /// <summary>
        /// Tries to parse a double from the specified <paramref name="candidateString"/> which is expected to be a string value.
        /// </summary>
        /// <param name="candidateString">The string value.</param>
        /// <param name="numericValue">The double value parsed from the specified <paramref name="candidateString"/>.</param>
        /// <param name="cultureInfo">Other <see cref="CultureInfo"/> than Current culture</param>
        /// <returns>True if <paramref name="candidateString"/> could be parsed to a double; otherwise, false.</returns>        
        internal static bool TryParseNumericString(string candidateString, out double numericValue, CultureInfo cultureInfo = null)
        {
            if (!string.IsNullOrEmpty(candidateString))
            {
                return double.TryParse(candidateString, ParseArguments.Number, cultureInfo ?? CultureInfo.CurrentCulture, out numericValue);
            }
            numericValue = 0;
            return false;
        }

        internal static bool TryParsePercentageString(string s, out double numericValue, CultureInfo cultureInfo = null)
        {
            numericValue = 0;
            if (string.IsNullOrEmpty(s)) return false;
            s = s.Trim();
            if (s.Trim().EndsWith("%"))
            {
                var tmp = s;
                while (tmp.Length > 0 && tmp.Last() == '%')
                {
                    tmp = tmp.Substring(0, tmp.Length - 1);
                }
                if (TryParseNumericString(tmp, out numericValue, cultureInfo))
                {
                    numericValue /= 100d;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Tries to parse a boolean value from the specificed <paramref name="candidateString"/>.
        /// </summary>
        /// <param name="candidateString">The value to check for boolean-ness.</param>
        /// <param name="result">The boolean value parsed from the specified <paramref name="candidateString"/>.</param>
        /// <returns>True if <paramref name="candidateString"/> could be parsed </returns>
        internal static bool TryParseBooleanString(string candidateString, out bool result)
        {
            if (!string.IsNullOrEmpty(candidateString))
            {
                if (candidateString == "-1" || candidateString == "1")
                {
                    result = true;
                    return true;
                }
                else if (candidateString == "0")
                {
                    result = false;
                    return true;
                }
                else
                {
                    return bool.TryParse(candidateString, out result);
                }
            }
            result = false;
            return false;
        }
        internal static bool ToBooleanString(string candidateString, bool defaultValue = false)
        {
            if (!string.IsNullOrEmpty(candidateString))
            {
                if (candidateString == "-1" || candidateString == "1")
                {
                    return true;
                }
                else if (candidateString == "0")
                {
                    return false;
                }
                else
                {
                    if (bool.TryParse(candidateString, out bool result))
                    {
                        return result;
                    }
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Tries to parse an int value from the specificed <paramref name="candidateString"/>.
        /// </summary>
        /// <param name="candidateString">The value to check for boolean-ness.</param>
        /// <param name="result">The boolean value parsed from the specified <paramref name="candidateString"/>.</param>
        /// <returns>True if <paramref name="candidateString"/> could be parsed </returns>
        internal static bool TryParseIntString(string candidateString, out int result)
        {
            if (!string.IsNullOrEmpty(candidateString))
                return int.TryParse(candidateString, out result);
            result = 0;
            return false;
        }
        internal static int? GetValueIntNull(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            return int.Parse(s, CultureInfo.InvariantCulture);
        }
        internal static long? GetValueLongNull(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            return long.Parse(s, CultureInfo.InvariantCulture);
        }
        /// <summary>
        /// Tries to parse a <see cref="DateTime"/> from the specified <paramref name="candidateString"/> which is expected to be a string value.
        /// </summary>
        /// <param name="candidateString">The string value.</param>
        /// <param name="result">The double value parsed from the specified <paramref name="candidateString"/>.</param>
        /// <returns>True if <paramref name="candidateString"/> could be parsed to a double; otherwise, false.</returns>
        internal static bool TryParseDateString(string candidateString, out DateTime result)
        {
            if (!string.IsNullOrEmpty(candidateString))
            {
                return DateTime.TryParse(candidateString, CultureInfo.CurrentCulture, ParseArguments.DateTime, out result);
            }
            result = DateTime.MinValue;
            return false;
        }
        internal static double GetValueDouble(object v, bool ignoreBool, bool retNaN, bool includeStrings)
        {
            double d;
            try
            {
                if (ignoreBool && v is bool)
                {
                    return 0;
                }
                if (IsNumericOrDate(v))
                {
                    if (v is DateTime)
                    {
                        d = ((DateTime)v).ToOADate();
                    }
                    else if (v is TimeSpan)
                    {
                        d = DateTime.FromOADate(0).Add((TimeSpan)v).ToOADate();
                    }
                    else
                    {
                        d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                    }
                }
                else if (includeStrings && v != null)
                {
                    if (TryParseNumericString(v.ToString(), out double val, CultureInfo.CurrentCulture))
                    {
                        d = val;
                    }
                    else if (TryParsePercentageString(v.ToString(), out double percentage, CultureInfo.CurrentCulture))
                    {
                        d = percentage;
                    }
                    else
                    {
                        d = retNaN ? double.NaN : 0;
                    }
                }
                else
                {
                    d = retNaN ? double.NaN : 0;
                }
            }

            catch
            {
                d = retNaN ? double.NaN : 0;
            }
            return d;
        }
        internal static string ExcelEscapeString(string s)
        {
            return s.Replace("&", "&amp;").
                     Replace("<", "&lt;").
                     Replace(">", "&gt;");
        }
        internal static string HtmlEscapeString(string s)
        {
            return s.Replace("&", "&amp;").
                     Replace("<", "&lt;").
                     Replace(">", "&gt;").
                     Replace("\"", "&quot;").
                     Replace("'", "&#39;");
        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sw"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StreamWriter sw, string t)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-Fa-f]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-Fa-f]{4,4}_)");
                int indexAdd = 0;
                while (match.Success)
                {
                    t = t.Insert(match.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                    match = match.NextMatch();
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && t[i] != '\t' && t[i] != '\n' && t[i] != '\r') //Not Tab, CR or LF
                {
                    sw.Write("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else if (t[i] > 0xFFFD)
                {
                    sw.Write($"_x{((int)t[i]).ToString("X")}_");
                }
                else
                {
                    sw.Write(t[i]);
                }
            }

        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="t"></param>
        /// <param name="encodeTabLF"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabLF = false)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-Fa-f]{4,4}_)"))
            {
                var matches = Regex.Matches(t, "(_x[0-9A-Fa-f]{4,4})");
                int indexAdd = 0;
                foreach (Match m in matches)
                {
                    t = t.Insert(m.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f &&
                    (t[i] != '\n' && t[i] != '\r' && t[i] != '\t' && encodeTabLF == false ||  //Not Tab, CR or LF
                    encodeTabLF))
                {
                    sb.AppendFormat("_x00{0}_", (t[i] <= 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else if (t[i] > 0xFFFD)
                {
                    sb.Append($"_x{((int)t[i]).ToString("X")}_");
                }
                else
                {
                    sb.Append(t[i]);
                }
            }
        }
        internal static string ExcelEscapeAndEncodeString(string t, bool crLfEncode = true)
        {
            if (string.IsNullOrEmpty(t))
            {
                return t;
            }
            return ExcelEncodeString(ExcelEscapeString(t), crLfEncode);
        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="t"></param>
        /// <param name="crLfEncode"></param>
        /// <returns></returns>
        internal static string ExcelEncodeString(string t, bool crLfEncode = true)
        {
            StringBuilder sb = new StringBuilder();
            t = t.Replace("\r\n", "\n"); //For some reason can't table name have cr in them. Replace with nl
            ExcelEncodeString(sb, t, crLfEncode);
            return sb.ToString();
        }
        internal static string ExcelDecodeString(string t)
        {
            if (string.IsNullOrEmpty(t)) return t;
            var ret = new StringBuilder();

            var ix = 0;
            for (var i = 0; i < t.Length; i++)
            {
                var c = t[i];
                if (c == '\r')
                {
                    ret.Append('\n');
                    if (i + 1 < t.Length && t[i + 1] == '\n')
                    {
                        i++;
                    }
                }
                else
                {
                    if (Matches(c, ref ix))
                    {
                        if (ix == 7)
                        {
                            var encoded = t.Substring(i - 4, 4);
                            ret.Append((char)int.Parse(encoded, NumberStyles.AllowHexSpecifier));
                            ix = 0;
                        }
                    }
                    else
                    {
                        if (ix > 0)
                        {
                            //If two underscores in a row the last one should count as potentially in matches
                            if (c == '_' && ix == 1)
                            {
                                ret.Append(t.Substring(i - ix, ix));
                            }
                            else
                            {
                            ret.Append(t.Substring(i - ix, ix + 1));
                                ix = 0;
                            }
                        }
                        else
                        {
                            ret.Append(c);
                        }
                    }
                }
            }

            if (ix > 0)
            {
                ret.Append(t.Substring(t.Length - ix, ix));
            }

            return ret.ToString();
        }

        private static bool Matches(char c, ref int ix)
        {
            if (ix == 0 && c == '_' ||
               ix == 1 && c == 'x' ||
               ix >= 2 && ix <= 5 && (c >= '0' && c <= '9' || c >= 'A' && c <= 'F' || c >= 'a' && c <= 'f') ||
               c == '_' && ix == 6)
            {
                ix++;
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool IsDate(object v, out DateTime date)
        {
            if (v is DateTime dt)
            {
                date = dt;
                return true;
            }
#if (NET6_0_OR_GREATER)
            if (v is DateOnly dto)
            {
                date = dto.ToDateTime(TimeOnly.MinValue);
                return true;
            }
#endif
            date = default;
            return false;
        }
        private static bool IsTimeSpan(object v, out TimeSpan timeSpan)
        {
            if (v is TimeSpan ts)
            {
                timeSpan = ts;
                return true;
            }
#if (NET6_0_OR_GREATER)
            if (v is TimeOnly to)
            {
                timeSpan = to.ToTimeSpan();
                return true;
            }
#endif
            timeSpan = default;
            return false;
        }
        /// <summary>
        /// Handles the issue with Excel having the incorrect values before 1900-03-01. Excel has 1900-02-29 as a valid value, but it does not exists in the calendar. Dates between 1900-02-28 and 1900-01-01 shuld be subtracted added 1 to the value. 0 in Excel is 1900-01-00 which is not valid in .NET. 
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        internal static DateTime FromOADateExcel(double d)
        {
            if (d < 0)
            {
                return DateTime.MinValue;
            }
            else if (d < 60)
            {
                d++;
            }
            return DateTime.FromOADate(d);
        }

        #region internal cache objects
        internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;
        internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.Name);  //TODO:Check that it works
        #endregion
    }
}
