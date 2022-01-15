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
    internal static class BinaryHelper
    {
        public static bool TryParseBinaryToDecimal(string number, int raiseToBase, out int result)
        {
            var chars = number.ToCharArray();
            var isNegative = chars[0] == '1';
            var negativeUsed = false;
            result = 0;
            for (var x = 1; x < 10; x++)
            {
                var c = chars[x];
                var current = 0;
                if (c != '0' && c != '1') return false;
                if (x == 9)
                {
                    current = c == '1' ? 1 : 0;
                    if (isNegative && !negativeUsed) current *= -1;
                }
                else if (c == '1')
                {
                    current = (int)System.Math.Pow(raiseToBase, 9 - x);
                    if (isNegative && !negativeUsed) current *= -1;
                    negativeUsed = true;
                }
                result += current;
            }
            return true;
        }
    }
}
