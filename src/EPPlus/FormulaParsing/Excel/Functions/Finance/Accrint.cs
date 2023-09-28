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
        EPPlusVersion = "6.0",
        Description = "Calculates the accrued interest for a security that pays periodic interest.")]
    internal class Accrint : ExcelFunction
    {
        public override int ArgumentMinLength => 6;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // collect input
            var id = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var issueDate = DateTime.FromOADate(id);

            var fid = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var firstInterestDate = DateTime.FromOADate(fid);

            var sd = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e1 != null) return CompileResult.GetErrorResult(e3.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var rate = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);

            var par = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if(e5 != null) return CompileResult.GetErrorResult(e5.Type);
            
            var frequency = ArgToInt(arguments, 5, out ExcelErrorValue e6);
            if(e6 != null) return CompileResult.GetErrorResult(e6.Type);
            var basis = 0;
            if(arguments.Count >= 7)
            {
                basis = ArgToInt(arguments, 6, out ExcelErrorValue e7);
                if (e7 != null) return CompileResult.GetErrorResult(e7.Type);
            }
            var issueToSettlement = true;
            if(arguments.Count >= 8)
            {
                issueToSettlement = ArgToBool(arguments, 7);
            }

            // validate input
            if (rate <= 0 || par <= 0) return CreateResult(eErrorType.Num);
            if (frequency != 1 && frequency != 2 && frequency != 4) return CreateResult(eErrorType.Num);
            if (basis < 0 || basis > 4) return CreateResult(eErrorType.Num);
            if (issueDate >= settlementDate) return CreateResult(eErrorType.Num);

            // calculation
            var dayCountBasis = (DayCountBasis)basis;
            var financialDays = FinancialDaysFactory.Create(dayCountBasis);
            var issue = FinancialDayFactory.Create(issueDate, dayCountBasis);
            var settlement = FinancialDayFactory.Create(settlementDate, dayCountBasis);
            var firstInterest = FinancialDayFactory.Create(firstInterestDate.AddDays(firstInterestDate.Day * -1 + 1), dayCountBasis);
            
            if(issueToSettlement)
            {
                var yearFrac = new YearFracProvider(context);
                var r = yearFrac.GetYearFrac(issueDate, settlementDate, dayCountBasis) * rate * par;
                return CreateResult(r, DataType.Decimal);
            }
            else
            {
                var r = CalculateInterest(issue, firstInterest, settlement, rate, par, frequency, dayCountBasis, context);
                return CreateResult(r, DataType.Decimal);
            }
        }

        private double CalculateInterest(FinancialDay issue, FinancialDay firstInterest, FinancialDay settlement, double rate, double par, int frequency, DayCountBasis basis, ParsingContext context)
        {
            var yearFrac = new YearFracProvider(context);
            var fds = FinancialDaysFactory.Create(basis);
            var nAdditionalPeriods = frequency == 1 ? 0 : 1;
            if(firstInterest <= settlement)
            {
                var p = fds.GetCalendarYearPeriodsBackwards(settlement, firstInterest, frequency, nAdditionalPeriods);
                var p2 = fds.GetCalendarYearPeriodsBackwards(firstInterest, settlement, frequency, nAdditionalPeriods);
                var firstPeriod = settlement >= firstInterest ? p.Last() : p.First();
                var yearFrac2 = yearFrac.GetYearFrac(firstPeriod.Start.ToDateTime(), settlement.ToDateTime(), basis);
                return yearFrac2 * rate * par;
            }
            else
            {
                var p2 = fds.GetCalendarYearPeriodsBackwards(firstInterest, settlement, frequency, nAdditionalPeriods);
                var firstInterestPeriod = p2.FirstOrDefault(x => x.Start < firstInterest && x.End >= firstInterest);
                var yearFrac2 = yearFrac.GetYearFrac(settlement.ToDateTime(), firstInterestPeriod.Start.ToDateTime(), basis) * -1;
                return yearFrac2 * rate * par;
            }
        }
    }
}
