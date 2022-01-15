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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class TwoComplementHelper
    {
        private static bool IsNegativeNumber(string candidate, int fromBase)
        {
            if (string.IsNullOrEmpty(candidate)) return false;
            return candidate.Length >= 10 && candidate.ToUpper().StartsWith(fromBase == 16 ? "F" : "7", StringComparison.OrdinalIgnoreCase);
        }

        public static double ParseDecFromString(string number, int fromBase)
        {
            if(IsNegativeNumber(number, fromBase))
            {
                return NegativeFromBase(number, fromBase) * -1;
            }
            return Convert.ToInt32(number, fromBase);
        }

        private static double NegativeFromBase(string number, int fromBase)
        {
            if (string.IsNullOrEmpty(number)) return double.NaN;
            var len = number.Length;
            var numArr = number.ToCharArray();
            var result = string.Empty;
            for (var x = len - 1; x >= 0; x--)
            {
                var part = Convert.ToInt32(numArr[x].ToString(), fromBase);
                if(fromBase == 16)
                {
                    result = (fromBase - 1 - part).ToString("X") + result;
                }
                else
                {
                    result = (fromBase - 1 - part).ToString() + result;
                }
            }
            var decResult = Convert.ToInt32(result, fromBase) + 1;
            return decResult;
        }
    }
}
