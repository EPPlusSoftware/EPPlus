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
        Description = "Calculates the bond-equivalent yield for a treasury bill")]
    internal class Tbilleq : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var discount = ArgToDecimal(arguments, 2);
            if (settlementDate >= maturityDate) return CompileResult.GetErrorResult(eErrorType.Num);
            if (discount <= 0d) return CompileResult.GetErrorResult(eErrorType.Num);
            var finDays = FinancialDaysFactory.Create(DayCountBasis.Actual_360);
            var nDaysInPeriod = finDays.GetDaysBetweenDates(settlementDate, maturityDate);
            if(nDaysInPeriod > 366)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            else if(nDaysInPeriod > 182)
            {
                var price = (100d - discount * 100d * nDaysInPeriod / 360d) / 100d;
                var fullYearDays = nDaysInPeriod <= 365 ? 365 : 366;
                var fullYearFactor = nDaysInPeriod / fullYearDays;
                var tmp = System.Math.Pow(fullYearFactor, 2) - (2d * fullYearFactor - 1d) * (1d - 1d / price);
                var term2 = System.Math.Sqrt(tmp);
                var term3 = 2d * fullYearFactor - 1d;
                var result = 2d * (term2 - fullYearFactor) / term3;
                return CreateResult(result, DataType.Decimal);
            }
            else
            {
                var result = (365d * discount) / (360d - (discount * nDaysInPeriod));
                return CreateResult(result, DataType.Decimal);
            }
        }
    }
}
