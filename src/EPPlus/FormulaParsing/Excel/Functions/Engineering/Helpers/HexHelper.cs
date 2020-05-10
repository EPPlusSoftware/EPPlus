/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Helpers
{
    internal static class HexHelper
    {
        private static bool IsNegativeHexNumber(string candidate)
        {
            if (string.IsNullOrEmpty(candidate)) return false;
            return candidate.Length >= 10 && candidate.ToUpper().StartsWith("F");
        }

        public static double GetDecFromHex(string number)
        {
            if(IsNegativeHexNumber(number))
            {
                return NegativeFromHex(number) * -1;
            }
            return Convert.ToInt32(number, 16);
        }

        private static double NegativeFromHex(string number)
        {
            if (string.IsNullOrEmpty(number)) return double.NaN;
            var len = number.Length;
            var numArr = number.ToCharArray();
            var result = string.Empty;
            for (var x = len - 1; x >= 0; x--)
            {
                var part = Convert.ToInt32(numArr[x].ToString(), 16);
                result = (15 - part).ToString("X") + result;
            }
            var decResult = Convert.ToInt32(result, 16) + 1;
            return decResult;
        }
    }
}
