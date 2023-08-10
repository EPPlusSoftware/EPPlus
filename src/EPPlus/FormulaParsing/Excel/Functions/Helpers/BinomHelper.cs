/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class BinomHelper
    {
        internal static double CumulativeDistrubution(double x, double trails, double probS)
        {
            var result = 0d;
            for (var i = 0; i <= x; i++)
            {
                var combin = MathHelper.Factorial(trails, trails - i) / MathHelper.Factorial(i);
                result += combin * Math.Pow(probS, i) * Math.Pow(1 - probS, trails - i);
            }
            return result;
        }
    }
}
