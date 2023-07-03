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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "6.0",
        Description = "Calculates the yield for a treasury bill")]
    internal class TbillYield : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = DateTime.FromOADate(ArgToInt(arguments, 1));
            var pr = ArgToDecimal(arguments, 2);
            if (settlementDate >= maturityDate) return CompileResult.GetErrorResult(eErrorType.Num);
            if (maturityDate.Subtract(settlementDate).TotalDays > 365) return CompileResult.GetErrorResult(eErrorType.Num);
            if (pr <= 0d) return CompileResult.GetErrorResult(eErrorType.Num);
            var finDays = FinancialDaysFactory.Create(DayCountBasis.Actual_360);
            var nDaysInPeriod = finDays.GetDaysBetweenDates(settlementDate, maturityDate);
            var result = ((100d - pr)/pr) * (360d/nDaysInPeriod);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
