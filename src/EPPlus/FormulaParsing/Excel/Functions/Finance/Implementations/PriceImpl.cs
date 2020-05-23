using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class PriceImpl
    {
        public static FinanceCalcResult<double> GetPrice(FinancialDay settlement, FinancialDay maturity, double rate, double yield, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360)
        {
            var coupDaysResult = new CoupdaysImpl(settlement, maturity, frequency, basis).GetCoupdays();
            if (coupDaysResult.HasError) return coupDaysResult;
            var coupdaysNcResult = new CoupdaysncImpl(settlement, maturity, frequency, basis).Coupdaysnc();
            if (coupdaysNcResult.HasError) return coupdaysNcResult;
            var coupnumResult = new CoupnumImpl(settlement, maturity, frequency, basis).GetCoupnum();
            if (coupnumResult.HasError) return new FinanceCalcResult<double>(coupnumResult.ExcelErrorType);
            var coupdaysbsResult = new CoupdaybsImpl(settlement, maturity, frequency, basis).Coupdaybs();
            if(coupdaysbsResult.HasError) return new FinanceCalcResult<double>(coupdaysbsResult.ExcelErrorType);

            var E = coupDaysResult.Result;
            var DSC = coupdaysNcResult.Result;
            var N = coupnumResult.Result;
            var A = coupdaysbsResult.Result;

            var retVal = -1d;
            if(N > 1)
            {
                var part1 = redemption / System.Math.Pow(1d + (yield / frequency), N - 1d + (DSC / E));
                var sum = 0d;
                for (var k = 1; k <= N; k++)
                {
                    sum += (100 * (rate / frequency)) / System.Math.Pow(1 + yield / frequency, k - 1 + DSC / E);
                }

                retVal = part1 + sum - (100 * (rate / frequency) * (A / E));
            }
            else
            {
                var DSR = E - A;
                var T1 = 100 * (rate / frequency) + redemption;
                var T2 = (yield / frequency) * (DSR / E) + 1;
                var T3 = 100 * (rate / frequency) * (A / E);

                retVal = T1 / T2 - T3;
            }

            return new FinanceCalcResult<double>(retVal);
        }
    }
}
