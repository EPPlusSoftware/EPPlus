/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
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
        EPPlusVersion = "5.5",
        Description = "Returns the annual yield of a security that pays interest at maturity.")]
    internal class Yieldmat : ExcelFunction
    {
        public override int ArgumentMinLength => 5;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = DateTime.FromOADate(ArgToInt(arguments, 1));
            if (settlementDate >= maturityDate) return CreateResult(eErrorType.Num);

            var issueDate = DateTime.FromOADate(ArgToInt(arguments, 2));
            
            var rate = ArgToDecimal(arguments, 3);
            if (rate < 0) return CompileResult.GetErrorResult(eErrorType.Num);
            
            var price = ArgToDecimal(arguments, 4);
            if (price <= 0) return CompileResult.GetErrorResult(eErrorType.Num);
            
            var basis = 0;
            if(arguments.Count > 5)
            {
                basis = ArgToInt(arguments, 5);
                if (basis < 0 || basis > 4) return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var yearFracProvider = new YearFracProvider(context);
            var yf1 = yearFracProvider.GetYearFrac(issueDate, maturityDate, (DayCountBasis)basis);
            var yf2 = yearFracProvider.GetYearFrac(issueDate, settlementDate, (DayCountBasis)basis);
            var yf3 = yearFracProvider.GetYearFrac(settlementDate, maturityDate, (DayCountBasis)basis);

            var result = 1d + yf1 * rate;
            result /= price / 100d + yf2 * rate;
            result = --result / yf3;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
