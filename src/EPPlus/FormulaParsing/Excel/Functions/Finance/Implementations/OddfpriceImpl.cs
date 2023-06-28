/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  26/06/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class OddfpriceImpl
    {
        private readonly System.DateTime _settlementDate;
        private readonly System.DateTime _maturityDate;
        private readonly System.DateTime _issueDate;
        private readonly System.DateTime _firstCouponDate;
        private readonly double _rate;
        private readonly double _yield;
        private readonly double _redemption;
        private readonly int _frequency;
        private readonly DayCountBasis _basis;

        public OddfpriceImpl(System.DateTime settlementDate, System.DateTime maturityDate, System.DateTime issueDate, System.DateTime firstCouponDate, double rate, double yield, double redemption, int frequency, DayCountBasis basis)
        {
            _settlementDate = settlementDate;
            _maturityDate = maturityDate;
            _issueDate = issueDate;
            _firstCouponDate = firstCouponDate;
            _rate = rate;
            _yield = yield;
            _redemption = redemption;
            _frequency = frequency;
            _basis = basis;
        }

        public FinanceCalcResult<double> GetOddfprice()
        {

            if (!((_maturityDate > _firstCouponDate)
               || (_maturityDate > _settlementDate)
               || (_maturityDate > _issueDate))
               || !((_firstCouponDate > _settlementDate)
               || (_firstCouponDate > _issueDate)
               || (_settlementDate > _issueDate)))
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if (_frequency != 1 && _frequency != 2 && _frequency != 4)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }
            var sDate = FinancialDayFactory.Create(_settlementDate, _basis);
            var mDate = FinancialDayFactory.Create(_maturityDate, _basis);
            var fcDate = FinancialDayFactory.Create(_firstCouponDate, _basis);
            var iDate = FinancialDayFactory.Create(_issueDate, _basis);
            var daysDefinition = FinancialDaysFactory.Create(_basis);

            // Many of the following variable names are taken from the formula in the Excel documentation for ODDFPRICE.
            // See https://support.microsoft.com/en-gb/office/oddfprice-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1

            var A = daysDefinition.GetDaysBetweenDates(iDate, sDate);
            var DSC = daysDefinition.GetDaysBetweenDates(sDate, fcDate);

            var coupDaysFunc = new CoupdaysImpl(sDate, fcDate, _frequency, _basis);
            var coupDaysResult = coupDaysFunc.GetCoupdays();
            var E = coupDaysResult.Result;

            var coupNumFunc = new CoupnumImpl(sDate, mDate, _frequency, _basis);
            var coupNumResult = coupNumFunc.GetCoupnum();
            var N = coupNumResult.Result;

            var DFC = daysDefinition.GetDaysBetweenDates(_issueDate, _firstCouponDate);
            var numOfMonths = 12 / _frequency;

            if (DFC < E)
            {
                // Short expression

                var t1 = _redemption / (System.Math.Pow(_yield / _frequency + 1, N - 1 + DSC / E));
                var t2 = (100 * _rate / _frequency * DFC / E) / (System.Math.Pow(1 + _yield / _frequency, DSC / E));

                var seriet3 = 0d;
                for (var i = 2; i <= N; i++)
                {
                    seriet3 += (100 * _rate / _frequency) / (System.Math.Pow(1 + _yield / _frequency, i - 1 + DSC / E));
                }

                var t3 = seriet3;

                var t4 = 100 * _rate / _frequency * A / E;

                var oddfprice_short = t1 + t2 + t3 - t4;

                return new FinanceCalcResult<double>(oddfprice_short);
            }
            else
            {
                // Long expression

                // Quasi periods: Normal period has to be divided into smaller period that match the _frequency
                // The interest in each quasi period is computed and the amounts are summed over the number of quasi
                // coupon periods.

                var coupNumfunc2 = new CoupnumImpl(iDate, fcDate, _frequency, _basis);
                var coupNumResult2 = coupNumfunc2.GetCoupnum();
                var NC = coupNumResult2.Result;

                // NC number of quasi periods in one odd period

                var quasiNumFunc = new CoupnumImpl(fcDate, sDate, _frequency, _basis);
                var quasiNumResult = quasiNumFunc.GetCoupnum();
                var Nq = quasiNumResult.Result;

                var coupNumFunc3 = new CoupnumImpl(fcDate, mDate, _frequency, _basis);
                var coupNumResult3 = coupNumFunc3.GetCoupnum();
                var N_long = coupNumResult3.Result;

                var lateCoup = _firstCouponDate;


                if (_basis == DayCountBasis.Actual_360 || _basis == DayCountBasis.Actual_365)
                {
                    var coupNcdFunc = new CoupncdImpl(sDate, fcDate, _frequency, _basis);
                    var coupNcdResult = coupNcdFunc.GetCoupncd();
                    var nextCoupDate = coupNcdResult.Result;
                    DSC = daysDefinition.GetDaysBetweenDates(_settlementDate, nextCoupDate);

                }
                else
                {
                    var coupPcdFunc = new CouppcdImpl(sDate, fcDate, _frequency, _basis);
                    var coupPcdResult = coupPcdFunc.GetCouppcd();
                    var previousCoupDate = coupPcdResult.Result;
                    A = daysDefinition.GetDaysBetweenDates(previousCoupDate, _settlementDate);
                    DSC = E - A;
                }

                var t1 = (_redemption) / (System.Math.Pow(1 + _yield / _frequency, N_long + Nq + DSC / E));

                var DCi = 0d;
                var NL = 0d;
                var dcDivNl = 0d;
                var aDivnl = 0d;
                var lateCouponDate = fcDate;
                var startDateDatetime = new System.DateTime(1900, 1, 1);
                var endDateDatetime = new System.DateTime(1900, 1, 1);

                var startDate = FinancialDayFactory.Create(startDateDatetime, _basis);
                var endDate = FinancialDayFactory.Create(endDateDatetime, _basis);


                for (var i = NC; i >= 1; i--)
                {

                    var earlyCouponDate = lateCouponDate.SubtractMonths(numOfMonths, lateCouponDate.Day);
                    if (_basis == DayCountBasis.Actual_Actual)
                    {

                        NL = daysDefinition.GetDaysBetweenDates(earlyCouponDate, lateCouponDate);
                    }
                    else
                    {
                        NL = E;
                    }

                    if (i > 1)
                    {
                        DCi = NL;
                    }
                    else
                    {
                        DCi = daysDefinition.GetDaysBetweenDates(iDate, lateCouponDate);
                    }

                    if (iDate > earlyCouponDate)
                    {
                        startDate = iDate;
                    }
                    else
                    {
                        startDate = earlyCouponDate;
                    }

                    if (sDate < lateCouponDate)
                    {
                        endDate = sDate;
                    }
                    else
                    {
                        endDate = lateCouponDate;
                    }

                    A = daysDefinition.GetDaysBetweenDates(startDate, endDate);
                    lateCouponDate = earlyCouponDate;

                    dcDivNl += DCi / NL;
                    aDivnl += A / NL;
                }

                var t2 = (100 * _rate / _frequency * dcDivNl) / (System.Math.Pow(1 + _yield / _frequency, Nq + DSC / E));
                var t3 = 0d;

                for (var k = 1; k <= N_long; k++)
                {
                    t3 += (100 * _rate / _frequency) / (System.Math.Pow(1 + _yield / _frequency, k - Nq + DSC / E));
                }

                var t4 = 100 * _rate / _frequency * aDivnl;

                var oddfprice_long = t1 + t2 + t3 - t4;

                return new FinanceCalcResult<double>(oddfprice_long);
            }
        }
    }
}
