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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal abstract class CoupFunctionBase<T> : ExcelFunction
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var sd = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var md = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturityDate = DateTime.FromOADate(md);

            var frequency = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var basis = 0;
            if (arguments.Count >= 4)
            {
                basis = ArgToInt(arguments, 3, out ExcelErrorValue e4);
                if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            }
            // validate input
            if((settlementDate > maturityDate) || (frequency != 1 && frequency != 2 && frequency != 4) || (basis < 0 || basis > 4))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            
            var result = ExecuteFunction(FinancialDayFactory.Create(settlementDate, (DayCountBasis)basis), FinancialDayFactory.Create(maturityDate, (DayCountBasis)basis), frequency, (DayCountBasis)basis);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            if (typeof(T) == typeof(DateTime))
            {
                return CreateResult(Convert.ToDateTime(result.Result).ToOADate(), DataType.Date);
            }
            return CreateResult(result.Result, result.DataType);
       }

        protected abstract FinanceCalcResult<T> ExecuteFunction(FinancialDay settlementDate, FinancialDay maturityDate, int frequency, DayCountBasis basis = DayCountBasis.US_30_360);
    }
}
