using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class IntRateImpl
    {
        internal static FinanceCalcResult<double> Intrate(DateTime settlement, DateTime maturity, double investment, double redemption, DayCountBasis basis = DayCountBasis.US_30_360)
        {
            if (investment <= 0 || redemption <= 0) return new FinanceCalcResult<double>(eErrorType.Num);
            if (maturity <= settlement) return new FinanceCalcResult<double>(eErrorType.Num);
            var settlementDay = FinancialDayFactory.Create(settlement, basis);
            var maturityDay = FinancialDayFactory.Create(maturity, basis);
            var fd = FinancialDaysFactory.Create(basis);
            var nDays = fd.GetDaysBetweenDates(settlementDay, maturityDay);
            
            // special case to make this function return same value as Excel
            if (basis == DayCountBasis.US_30_360 && maturityDay.Day == 31) nDays++;

            var result = ((redemption - investment) / investment) * fd.DaysPerYear / nDays;
            return new FinanceCalcResult<double>(result);
        }
    }
}
