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
    }
}
