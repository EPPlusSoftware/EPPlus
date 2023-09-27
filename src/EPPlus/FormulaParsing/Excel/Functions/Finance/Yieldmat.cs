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
            var sd = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var md = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturityDate = DateTime.FromOADate(md);
            
            if (settlementDate >= maturityDate) return CreateResult(eErrorType.Num);

            var id = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var issueDate = DateTime.FromOADate(id);
            
            var rate = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            
            if (rate < 0) return CompileResult.GetErrorResult(eErrorType.Num);
            
            var price = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            if (price <= 0) return CompileResult.GetErrorResult(eErrorType.Num);
            
            var basis = 0;
            if(arguments.Count > 5)
            {
                basis = ArgToInt(arguments, 5, out ExcelErrorValue e6);
                if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
                
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
