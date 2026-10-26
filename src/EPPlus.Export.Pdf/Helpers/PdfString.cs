/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Globalization;

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class PdfString
    {
        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string.</returns>
        internal static string ToPdfString(this double val)
        {
            return val.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string with 4 decimals.</returns>
        internal static string ToPdfStringF4(this double val)
        {
            return val.ToString("F4", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string with no decimals.</returns>
        internal static string ToPdfStringF0(this double val)
        {
            return val.ToString("F0", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string with no decimals.</returns>
        internal static string ToPdfStringF0(this int val)
        {
            return val.ToString("F0", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string with no decimals.</returns>
        internal static string ToPdfStringF0(this int? val)
        {
            int v = val ?? 0;
            if (val == null)
                return string.Empty;
            return v.ToString("F0", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value formated for use in pdf document.
        /// </summary>
        /// <param name="val">Value to turn into a string.</param>
        /// <returns>The value repsented as a string with no decimals.</returns>
        internal static string ToPdfStringF0(this long val)
        {
            return val.ToString("F0", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the value escaped and wrapped in parentheses as a PDF literal string (escaping '(', ')' and '\'), for use in places such as a link annotation's URI.
        /// </summary>
        /// <param name="val">The text to encode, for example a hyperlink URI.</param>
        /// <returns>The value as an escaped PDF literal string.</returns>
        internal static string ToPdfLiteral(this string s)
        {
            if (s == null) return "()";
            var sb = new System.Text.StringBuilder(s.Length + 2);
            sb.Append('(');
            foreach (char c in s)
            {
                if (c == '(' || c == ')' || c == '\\') sb.Append('\\');
                if (c == '\r') { sb.Append("\\r"); continue; }
                if (c == '\n') { sb.Append("\\n"); continue; }
                sb.Append(c);
            }
            sb.Append(')');
            return sb.ToString();
        }

        internal static string ToPdfTextString(this string val)
        {
            if (string.IsNullOrEmpty(val)) return "()";
            var sb = new System.Text.StringBuilder(val.Length * 4 + 6);
            sb.Append("<FEFF");
            foreach (char c in val) sb.Append(((int)c).ToString("X4"));  // char = UTF-16 code unit
            sb.Append('>');
            return sb.ToString();
        }
    }
}
