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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Returns the Fisher transformation at x. This transformation produces a function that is normally distributed rather than skewed.")]
    internal class Geomean : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var numbers = ArgsToDoubleEnumerable(arguments, context);
            if (numbers.Any(x => x.Value <= 0d)) return CreateResult(eErrorType.Num);
            var p = 1d;
            for(var x = 0; x < numbers.Count(); x++)
            {
                var n = numbers.ElementAt(x);
                p *= n.Value;
            }
            var result = System.Math.Pow(p, 1d / numbers.Count());
            return CreateResult(result, DataType.Decimal);
        }
    }
}
