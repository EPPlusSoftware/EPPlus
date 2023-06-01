/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/03/2020         EPPlus Software AB           EPPlus 5.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "5.1",
        Description = "Returns the sum of the difference of squares of corresponding values in two supplied arrays")]
    internal class SumX2mY2 : SumxBase
    {
        public override int ArgumentMinLength => 2;
        public override double Calculate(double[] set1, double[] set2)
        {
            var result = 0d;
            for(var x = 0; x < set1.Length; x++)
            {
                var a = set1[x];
                var b = set2[x];
                result += a * a - b * b;
            }
            return result;
        }
    }
}
