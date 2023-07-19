/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/07/2023         EPPlus Software AB           EPPlus v7
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
    EPPlusVersion = "7.0",
    Description = "Returns the Poisson distribution. This function works the same as POISSON.DIST")]
    internal class Poisson : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0);
            var mean = ArgToDecimal(arguments, 1);
            var cumulative = ArgToBool(arguments, 2);

            x = Math.Floor(x);

            if (x < 0 || mean < 0)
            {
                return CreateResult(eErrorType.Num);
            }

            if (cumulative)
            {
                var meanDivFac = 0d;
                for (var k = 0; k <= x; k++)
                {
                    var facResult = MathHelper.Factorial(k);
                    var meanK = Math.Pow(mean, k);
                    meanDivFac += meanK / facResult;
                }

                var result = Math.Exp(-mean) * meanDivFac;

                return CreateResult(result, DataType.Decimal);
            }
            else
            {
                return CreateResult(Math.Pow(mean, x) * Math.Exp(-mean) / MathHelper.Factorial(x), DataType.Decimal);
            }
        }

    }
}
