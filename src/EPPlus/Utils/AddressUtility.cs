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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// A utility to work with Excel addresses
    /// </summary>
    public static class AddressUtility
    {
        /// <summary>
        /// Parse an entire column selection, e.g A:A
        /// </summary>
        /// <param name="address">The entire address</param>
        /// <returns></returns>
        public static string ParseEntireColumnSelections(string address)
        {
            string parsedAddress = address;
            var matches = Regex.Matches(address, "[A-Z]+:[A-Z]+");
            foreach (Match match in matches)
            {
                AddRowNumbersToEntireColumnRange(ref parsedAddress, match.Value);
            }
            return parsedAddress;
        }
        /// <summary>
        /// Add row number to entire column range
        /// </summary>
        /// <param name="address">The address</param>
        /// <param name="range">The full column range</param>
        private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
        {
            var parsedRange = string.Format("{0}{1}", range, ExcelPackage.MaxRows);
            var splitArr = parsedRange.Split(new char[] { ':' });
            address = address.Replace(range, string.Format("{0}1:{1}", splitArr[0], splitArr[1]));
        }
    }
}
