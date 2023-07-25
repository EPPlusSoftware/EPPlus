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
        Description = "Returns the inverse of the lognormal cumulative distribution function")]
    internal class LognormDotInv : NormalDistributionBase
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var p = ArgToDecimal(arguments, 0);
            var mean = ArgToDecimal(arguments, 1);
            var stdev = ArgToDecimal(arguments, 2);
            if (p <= 0|| p>=1||stdev<=0)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var result = Math.Exp(-Math.Sqrt(2*ErfHelper.Erfcinv(2*p)));
            return CreateResult(result, DataType.Decimal);
        }
    }
}