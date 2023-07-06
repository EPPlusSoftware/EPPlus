/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the price per $100 face value of a security that pays periodic interest")]
    internal class Price : ExcelFunction
    {
        public override int ArgumentMinLength => 6;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = DateTime.FromOADate(ArgToInt(arguments, 1));
            var rate = ArgToDecimal(arguments, 2);
            var yield = ArgToDecimal(arguments, 3);
            var redemption = ArgToDecimal(arguments, 4);
            var frequency = ArgToInt(arguments, 5);
            var basis = 0;
            if (arguments.Count >= 7)
            {
                basis = ArgToInt(arguments, 6);
            }
            // validate input
            if ((settlementDate > maturityDate) || rate < 0 || yield < 0 || redemption <= 0 || (frequency != 1 && frequency != 2 && frequency != 4) || (basis < 0 || basis > 4))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var result = PriceImpl.GetPrice(FinancialDayFactory.Create(settlementDate, (DayCountBasis)basis), FinancialDayFactory.Create(maturityDate, (DayCountBasis)basis), rate, yield, redemption, frequency, (DayCountBasis)basis);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
