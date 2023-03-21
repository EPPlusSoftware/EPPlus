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
        Description = "Calculates the price per $100 face value for a treasury bill")]
    internal class TbillPrice : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var discount = ArgToDecimal(arguments, 2);
            if (settlementDate >= maturityDate) return CreateResult(eErrorType.Num);
            if (maturityDate.Subtract(settlementDate).TotalDays > 365) return CreateResult(eErrorType.Num);
            if (discount <= 0d) return CreateResult(eErrorType.Num);
            var finDays = FinancialDaysFactory.Create(DayCountBasis.Actual_360);
            var nDaysInPeriod = finDays.GetDaysBetweenDates(settlementDate, maturityDate);
            var result = 100d * (1d - (discount * nDaysInPeriod / 360d));
            return CreateResult(result, DataType.Decimal);
        }
    }
}
