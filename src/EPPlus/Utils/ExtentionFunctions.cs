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
using System.Linq;

namespace OfficeOpenXml.Utils.Extentions
{
    internal static class EnumExtensions
    {
        /// <summary>
        /// Returns the enum value with first char lower case
        /// </summary>
        /// <param name="enumValue"></param>
        /// <returns></returns>
        internal static string ToEnumString(this Enum enumValue)
        {
            var s = enumValue.ToString();
            return s.Substring(0, 1).ToLower() + s.Substring(1);
        }
        internal static T ToEnum<T>(this string s, T defaultValue) where T : struct
        {
            try
            {
                if (string.IsNullOrEmpty(s)) return defaultValue;
                return (T)Enum.Parse(typeof(T), s, true);
            }
            catch
            {                
                return defaultValue;
            }
        }

        internal static string GetStringValueForXml(this bool boolValue)
        {
            return boolValue ? "1" : "0";
        }
        internal static bool IsInt(this string s)
        {
            return (!s.Any(x => x < '0' && x > '9'));
        }
    }
}
