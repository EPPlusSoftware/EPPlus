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
        Description = "Calculates the cumulative beta probability density function")]
    internal class Betadist : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override string NamespacePrefix => "_xlfn.";
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0);
            var alpha = ArgToDecimal(arguments, 1);
            var beta = ArgToDecimal(arguments, 2);
            var A = 0d;
            var B = 1d;
            if (arguments.Count() > 3)
            {
                A = ArgToDecimal(arguments, 3);
            }
            if (arguments.Count() > 4)
            {
                B = ArgToDecimal(arguments, 4);
            }
            // validate
            if (alpha <= 0 || beta <= 0) return CreateResult(eErrorType.Num);
            if (x < A || x > B || A == B) return CreateResult(eErrorType.Num);
            x = (x - A) / (B - A);
            var result = BetaHelper.BetaCdf(x, alpha, beta);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
