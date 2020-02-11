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
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.IO;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Utils
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
            return (TypeCompat.IsPrimitive(candidate) || candidate is double || candidate is decimal || candidate is DateTime || candidate is TimeSpan || candidate is long);
        }
        internal static bool IsNumeric(object candidate)
        {
            if (candidate == null) return false;
            return (TypeCompat.IsPrimitive(candidate) || candidate is double || candidate is decimal || candidate is long);
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
        /// <summary>
        /// Tries to parse a boolean value from the specificed <paramref name="candidateString"/>.
        /// </summary>
        /// <param name="candidateString">The value to check for boolean-ness.</param>
        /// <param name="result">The boolean value parsed from the specified <paramref name="candidateString"/>.</param>
        /// <returns>True if <paramref name="candidateString"/> could be parsed </returns>
        internal static bool TryParseBooleanString(string candidateString, out bool result)
		{
			if (!string.IsNullOrEmpty(candidateString))
				return bool.TryParse(candidateString, out result);
			result = false;
			return false;
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
		/// <summary>
		/// Convert an object value to a double 
		/// </summary>
		/// <param name="v"></param>
		/// <param name="ignoreBool"></param>
        /// <param name="retNaN">Return NaN if invalid double otherwise 0</param>
		/// <returns></returns>
		internal static double GetValueDouble(object v, bool ignoreBool = false, bool retNaN=false)
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
        internal static DateTime? GetValueDate(object v)
        {
            if (v is DateTime d)
            {
                return d;
            }
            else
            {
                try
                {
                    if (IsNumericOrDate(v))
                    {
                        var n = GetValueDouble(v);
                        if (double.IsNaN(n))
                        {
                            return DateTime.FromOADate(n);
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }
        internal static string ExcelEscapeString(string s)
        {
            return s.Replace("&", "&amp;").
                     Replace("<", "&lt;").
                     Replace(">", "&gt;");
        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sw"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StreamWriter sw, string t)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
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
        /// <param name="encodeTabCRLF"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabCRLF=false)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
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
                if (t[i] <= 0x1f && ((t[i] != '\t' && t[i] != '\n' && t[i] != '\r' && encodeTabCRLF == false) || encodeTabCRLF)) //Not Tab, CR or LF
                {
                    sb.AppendFormat("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else
                {
                    sb.Append(t[i]);
                }
            }

        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static string ExcelEncodeString(string t)
        {
            StringBuilder sb=new StringBuilder();
            t=t.Replace("\r\n", "\n"); //For some reason can't table name have cr in them. Replace with nl
            ExcelEncodeString(sb, t, true);
            return sb.ToString();
        }
        internal static string ExcelDecodeString(string t)
        {
            var match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
            if (!match.Success) return t;

            var useNextValue = false;
            var ret = new StringBuilder();
            var prevIndex = 0;
            while (match.Success)
            {
                if (prevIndex < match.Index) ret.Append(t.Substring(prevIndex, match.Index - prevIndex));
                if (!useNextValue && match.Value == "_x005F")
                {
                    useNextValue = true;
                }
                else
                {
                    if (useNextValue)
                    {
                        ret.Append(match.Value);
                        useNextValue = false;
                    }
                    else
                    {
                        ret.Append((char)int.Parse(match.Value.Substring(2, 4), NumberStyles.AllowHexSpecifier));
                    }
                }
                prevIndex = match.Index + match.Length;
                match = match.NextMatch();
            }
            ret.Append(t.Substring(prevIndex, t.Length - prevIndex));
            return ret.ToString();
        }

        /// <summary>
        ///     Convert cell value to desired type, including nullable structs.
        ///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
        ///     When attempted conversion fails exception is passed through.
        /// </summary>
        /// <typeparam name="T">
        ///     The type to convert to.
        /// </typeparam>
        /// <returns>
        ///     The <paramref name="value"/> converted to <typeparamref name="T"/>.
        /// </returns>
        /// <remarks>
        ///     If input is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
        ///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
        ///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
        ///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
        /// </remarks>
        /// <exception cref="FormatException">
        ///     <paramref name="value"/> is string and its format is invalid for conversion (parsing fails)
        /// </exception>
        /// <exception cref="InvalidCastException">
        ///     <paramref name="value"/> is not string and direct conversion fails
        /// </exception>
        public static T GetTypedCellValue<T>(object value)
        {
            var conversion = new TypeConvertUtil<T>(value);
            if(value == null || (conversion.ReturnType.IsNullable && conversion.Value.IsEmptyString))
            {
                return default;
            }
            else if (value.GetType() == conversion.ReturnType.Type)
            {
                return (T)value;
            }
            else if ((conversion.Value.IsString || conversion.Value.IsNumeric) && conversion.ReturnType.IsNumeric)
            {
                return (T)conversion.ConvertToReturnType();
            }
            else if (conversion.ReturnType.IsDateTime && conversion.TryGetDateTime(out object returnDate))
            {
                return (T)returnDate;
            }
            else if (conversion.ReturnType.IsTimeSpan && conversion.TryGetTimeSpan(out object ts))
            {
                return (T)ts;
            }

            return (T)Convert.ChangeType(value, conversion.ReturnType.Type);
        }
        internal static string GetValueForXml(object v, bool date1904)
        {
            string s;
            try
            {
                if (v is DateTime)
                {
                    double sdv = ((DateTime)v).ToOADate();

                    if(date1904)
                    {
                        sdv -= ExcelWorkbook.date1904Offset;
                    }

                    s = sdv.ToString(CultureInfo.InvariantCulture);
                }
                else if (v is TimeSpan)
                {
                    s = DateTime.FromOADate(0).Add(((TimeSpan)v)).ToOADate().ToString(CultureInfo.InvariantCulture);
                }
                else if (TypeCompat.IsPrimitive(v) || v is double || v is decimal)
                {
                    if ((v is double && double.IsNaN((double)v)) ||
                        (v is float && float.IsNaN((float)v)))
                    {
                        s = "";
                    }
                    else if (v is double && double.IsInfinity((double)v))
                    {
                        s = "#NUM!";
                    }
                    else
                    {
                        s = Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    s = v.ToString();
                }
            }

            catch
            {
                s = "0";
            }
            return s;
        }
        #region internal cache objects
        internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;
        internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.Name);  //TODO:Check that it works
        #endregion
    }
}
