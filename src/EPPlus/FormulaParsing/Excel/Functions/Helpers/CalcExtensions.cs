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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class CalcExtensions
    {
        public static double AverageKahan(this IEnumerable<double> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val;
            }
            return sum.Get() / values.Count();
        }

        public static double AverageKahan(this IEnumerable<int> values, Func<int, double> selector)
        {
            KahanSum sum = 0.0;
            foreach(var val in values)
            {
                sum += selector.Invoke(val);
            }
            return sum.Get() / values.Count();
        }
    }
}
