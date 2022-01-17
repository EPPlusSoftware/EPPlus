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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
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
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 6);
            // collect input
            var issueDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var firstInterestDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 2));
            var rate = ArgToDecimal(arguments, 3);
            var par = ArgToDecimal(arguments, 4);
            var frequency = ArgToInt(arguments, 5);
            var basis = 1;
            if(arguments.Count() >= 7)
            {
                basis = ArgToInt(arguments, 6);
            }
            var issueToSettlement = true;
            if(arguments.Count() >= 8)
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
            
            var oddPeriodStart = issueToSettlement ? issue : firstInterest;

            var coupNumFunc = new CoupnumImpl(oddPeriodStart, settlement, frequency, dayCountBasis);
            var nc = coupNumFunc.GetCoupnum().Result;
            var result = par * (rate / frequency);
            var lastPart = 0d;
            var startDate = GetCoupPcd(oddPeriodStart, settlement, frequency, dayCountBasis);
            var endDate = GetCoupNcd(oddPeriodStart, settlement, frequency, dayCountBasis);

            // create a "fake" coupon period
            var fds = FinancialDaysFactory.Create(dayCountBasis);
            var periods = default(FinancialPeriod[]);
            if(issueToSettlement)
            {
                periods = fds.GetCouponPeriodsBackwards(settlement, oddPeriodStart, frequency).ToArray();
            }
            else
            {
                periods = fds.GetCalendarYearPeriodsBackwards(settlement, oddPeriodStart, frequency).ToArray();
            }

            foreach(var period in periods)
            {
                var periodDays = fds.GetDaysBetweenDates(period.Start, period.End);
                var nDaysInPeriod = fds.GetDaysBetweenDates(period.Start, period.End);
                if(period.Start < oddPeriodStart && issueToSettlement)
                {
                    nDaysInPeriod = oddPeriodStart.SubtractDays(period.End);
                }
                else if(period.End > settlement)
                {
                    nDaysInPeriod = periodDays - System.Math.Abs(period.End.SubtractDays(settlement));
                }
                // för första perioden: om couponperiods start datum < settlement. Räkna med antal Financial days från settlement till periodens slut.
                // för sista perioden: om couponperiods slutdatum > firstInterest. Räkna med antal Financial days från sista periodens start till first interest.
                // annars:
                // använd Coupdays för perioden.
                

                lastPart += nDaysInPeriod / periodDays;
            }
            result *= lastPart + (issueToSettlement ? 0 : 1);
            return CreateResult(result, DataType.Decimal);
        }

        private FinancialPeriod GetPreviousCouponPeriod(System.DateTime currentPeriodStartDate, int frequency, DayCountBasis basis)
        {
            var periodEndDate = FinancialDayFactory.Create(currentPeriodStartDate.AddDays(-1), basis);
            var fds = FinancialDaysFactory.Create(basis);
            var result = fds.GetCouponPeriod(periodEndDate, FinancialDayFactory.Create(currentPeriodStartDate, basis), frequency);
            return result;
        }

        private System.DateTime GetCoupNcd(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis dayCountBasis)
        {
            var coupNcdFunc = new CoupncdImpl(settlement, maturity, frequency, dayCountBasis);
            return coupNcdFunc.GetCoupncd().Result;
        }

        private System.DateTime GetCoupPcd(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis dayCountBasis)
        {
            var coupNcdFunc = new CouppcdImpl(settlement, maturity, frequency, dayCountBasis);
            return coupNcdFunc.GetCouppcd().Result;
        }

        private double GetCoupdays(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis dayCountBasis)
        {
            var coupDaysFunc = new CoupdaysImpl(settlement, maturity, frequency, dayCountBasis);
            return coupDaysFunc.GetCoupdays().Result;
        }
    }
}
