/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Returns the Pearson product moment correlation coefficient.")]
    internal class Pearson : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            var array1 = ArgsToDoubleEnumerable(arg1, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var array2 = ArgsToDoubleEnumerable(arg2, context, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (array1.Count != array2.Count) return CompileResult.GetErrorResult(eErrorType.NA);
            if (!array1.Any()) return CompileResult.GetErrorResult(eErrorType.NA);
            var result = PearsonImpl(array1, array2);
            return CreateResult(result, DataType.Decimal);
        }

        internal static double PearsonImpl(IEnumerable<double> arr1, IEnumerable<double> arr2)
        {
            var avg1 = arr1.AverageKahan();
            var avg2 = arr2.AverageKahan();
            var length = arr1.Count();
            var number = 0d;
            double d1 = 0d, d2 = 0d;
            for(var x = 0; x < length; x++)
            {
                number += (arr1.ElementAt(x) - avg1) * (arr2.ElementAt(x) - avg2);
                d1 += Math.Pow(arr1.ElementAt(x) - avg1, 2);
                d2 += Math.Pow(arr2.ElementAt(x) - avg2, 2);
            }
            return number / Math.Sqrt(d1 * d2);
        }
    }
}
