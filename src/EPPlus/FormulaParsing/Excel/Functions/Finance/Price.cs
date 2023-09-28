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
            var sd = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var md = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturityDate = DateTime.FromOADate(md);

            var rate = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            
            var yield = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            
            var redemption = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            
            var frequency = ArgToInt(arguments, 5, out ExcelErrorValue e6);
            if(e6 != null) return CompileResult.GetErrorResult(e6.Type);

            var basis = 0;
            if (arguments.Count >= 7)
            {
                basis = ArgToInt(arguments, 6, out ExcelErrorValue e7);
                if (e7 != null) return CompileResult.GetErrorResult(e7.Type);
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
