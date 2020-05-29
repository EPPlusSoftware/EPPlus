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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "5.1",
        Description = "Returns the sum of a power series")]
    internal class Seriessum : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            var x = ArgToDecimal(arguments, 0);
            var n = ArgToDecimal(arguments, 1);
            var m = ArgToDecimal(arguments, 2);
            var coeffs = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(3) }, context).ToArray();
            var result = 0d;
            for(var i = 0; i < coeffs.Count(); i++)
            {
                var c = coeffs[i];
                result += c * System.Math.Pow(x, (i * n) + m);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
