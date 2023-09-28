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
        Description = "Calculates the inverse of the beta cumulative probability density function")]
    internal class Betainv : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var probability = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var alpha = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            var beta = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var A = 0d;
            var B = 1d;
            if(arguments.Count > 3)
            {
                A = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
                if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            }
            if (arguments.Count > 4)
            {
                B = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
                if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            }
            // validate
            if (alpha <= 0 || beta <= 0) return CreateResult(eErrorType.Num);
            if (probability <= 0 || probability > 1) return CreateResult(eErrorType.Num);

            var result = BetaHelper.IBetaInv(probability, alpha, beta) * (B - A) + A;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
