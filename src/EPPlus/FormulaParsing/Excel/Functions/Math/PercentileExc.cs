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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
     Category = ExcelFunctionCategory.Statistical,
     EPPlusVersion = "5.5",
     IntroducedInExcelVersion = "2010",
     Description = "Returns the K'th percentile of values in a supplied range, where K is in the range 0 - 1 (exclusive)")]
    internal class PercentileExc : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arr = ArgsToDoubleEnumerable(arguments.Take(1), context).Select(x => (double)x).ToList();
            var k = ArgToDecimal(arguments, 1);
            if (k <= 0 || k >= 1) return CreateResult(eErrorType.Num);
            var n = arr.Count;
            if (k < 1d / (n + 1d) || k > 1 - 1d / (n + 1d)) return CreateResult(eErrorType.Num);
            arr.Sort();
            var l = k * (n + 1d) - 1;
            var fl = (int)System.Math.Floor(l);
            var result = ((l - fl) < double.Epsilon) ? arr[fl] : arr[fl] + (l - fl) * (arr[fl + 1] - arr[fl]);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
