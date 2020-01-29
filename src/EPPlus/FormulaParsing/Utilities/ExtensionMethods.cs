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
using OfficeOpenXml.Compatibility;
using System;
namespace OfficeOpenXml.FormulaParsing.Utilities
{
    internal static class ExtensionMethods
    {
        internal static void IsNotNullOrEmpty(this ArgumentInfo<string> val)
        {
            if (string.IsNullOrEmpty(val.Value))
            {
                throw new ArgumentException(val.Name + " cannot be null or empty");
            }
        }

        internal static void IsNotNull<T>(this ArgumentInfo<T> val)
            where T : class
        {
            if (val.Value == null)
            {
                throw new ArgumentNullException(val.Name);
            }
        }

        internal static bool IsNumeric(this object obj)
        {
            if (obj == null) return false;
            return (TypeCompat.IsPrimitive(obj) || obj is double || obj is decimal || obj is System.DateTime || obj is TimeSpan);
        }
    }
}
