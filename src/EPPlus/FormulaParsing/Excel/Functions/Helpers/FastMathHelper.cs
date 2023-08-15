/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class FastMathHelper
    {
        internal static double Min(params double[] values)
        {
            var minValue = double.MaxValue;
            foreach (var d in values)
            {
                if (d < minValue)
                {
                    minValue = d;
                }
            }
            return minValue;
        }

        internal static int Min(params int[] values)
        {
            var minValue = int.MaxValue;
            foreach (var i in values)
            {
                if (i < minValue)
                {
                    minValue = i;
                }
            }
            return minValue;
        }
    }
}
