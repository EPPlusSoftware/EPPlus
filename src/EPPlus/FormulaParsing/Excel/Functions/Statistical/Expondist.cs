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
            Description = "Returns the value of the exponential distribution for a give value of x. Same implementation as EXPON.DIST")]
    internal class Expondist : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var x = ArgToDecimal(arguments, 0);
            var lambda = ArgToDecimal(arguments, 1);
            var cumulative = ArgToBool(arguments, 2);
            if (lambda <= 0d) return CreateResult(eErrorType.Num);
            if (x < 0d) return CreateResult(eErrorType.Num);
            var result = 0d;
            if (cumulative && x >= 0)
            {
                result = 1d - System.Math.Exp(x * -lambda);
            }
            else if(!cumulative && x >= 0)
            {
                result = lambda * System.Math.Exp(x * -lambda);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
