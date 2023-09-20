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
        Description = "Calculates the cumulative principal paid on a loan, between two specified periods")]
    internal class Cumprinc : ExcelFunction
    {
        public override int ArgumentMinLength => 6;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rate = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var nPer = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            var presentValue = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if(e3 != null) return CreateResult(e3.Type);
            var startPeriod = ArgToInt(arguments, 3);
            var endPeriod = ArgToInt(arguments, 4);
            var t = ArgToInt(arguments, 5);
            if (t < 0 || t > 1) return CompileResult.GetErrorResult(eErrorType.Num);
            var func = new CumprincImpl(new PmtProvider(), new FvProvider());
            var result = func.GetCumprinc(rate, nPer, presentValue, startPeriod, endPeriod, (PmtDue)t);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }

        private static double GetInterest(double rate, double remainingAmount)
        {
            return remainingAmount * rate;
        }
    }
}
