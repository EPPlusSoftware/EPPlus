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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
        Description = "Calculates the kurtosis of a data set")]
    internal class Kurt : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var numbers = ArgsToDoubleEnumerable(arguments, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var n = (double)numbers.Count();
            if (n < 4) return CompileResult.GetErrorResult(eErrorType.Div0);
            var stdev = new Stdev().StandardDeviation(numbers);
            if(stdev.DataType == DataType.ExcelError)
            {
                return stdev;
            }
            var part1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3));
            var avg = numbers.Average();
            var part2 = 0d;
            for(var x = 0; x < n; x++)
            {
                part2 += System.Math.Pow((numbers.ElementAt(x) - avg), 4);
            }
            part2 /= System.Math.Pow((double)stdev.Result, 4);
            var result = part1 * part2 - (3 * System.Math.Pow(n - 1, 2)) / ((n - 2) * (n - 3));
            return CreateResult(result, DataType.Decimal);
        }
    }
}
