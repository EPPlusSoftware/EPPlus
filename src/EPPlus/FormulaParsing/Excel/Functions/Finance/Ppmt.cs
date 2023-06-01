/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the payments required to reduce a loan, from a supplied present value to a specified future value")]
    internal class Ppmt : ExcelFunction
    {
        public override int ArgumentMinLength => 4;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rate = ArgToDecimal(arguments, 0);
            var per = ArgToInt(arguments, 1);
            var nPer = ArgToInt(arguments, 2);
            var presentValue = ArgToDecimal(arguments, 3);
            var fv = 0d;
            if (arguments.Count >= 5)
            {
                fv = ArgToDecimal(arguments, 4);
            }
            var type = PmtDue.EndOfPeriod;
            if (arguments.Count >= 6)
            {
                type = (PmtDue)ArgToInt(arguments, 5);
            }
            var result = PpmtImpl.Ppmt(rate, per, nPer, presentValue, fv, type);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
