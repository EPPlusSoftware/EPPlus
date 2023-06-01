/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "5.2",
    Description = "Returns the K'th percentile of values in a supplied range, where K is in the range 0 - 1 (inclusive)")]
    internal class PercentileInc : HiddenValuesHandlingFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arr = ArgsToDoubleEnumerable(arguments.Take(1), context).Select(x => (double)x).ToList();
            var percentile = ArgToDecimal(arguments, 1);
            if (percentile < 0 || percentile > 1) return CompileResult.GetErrorResult(eErrorType.Num);
            arr.Sort();
            var nElements = arr.Count;
            var dIx = percentile * (nElements - 1);
            var ix = (int)dIx;
            var rest = dIx - ix;
            var result = ix < (nElements - 1) ? arr[ix] + (arr[ix + 1] - arr[ix]) * rest : arr.Last();
            return CreateResult(result, DataType.Decimal);
        }
    }
}
