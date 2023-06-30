/*************************************************************************************************
* This Source Code Form is subject to the terms of the Mozilla Public
* License, v. 2.0. If a copy of the MPL was not distributed with this
* file, You can obtain one at http://mozilla.org/MPL/2.0/.
*************************************************************************************************
Date               Author                       Change
*************************************************************************************************
05/20/2020         EPPlus Software AB       Implemented function
*************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class DurationImpl
    {
        public DurationImpl(IYearFracProvider yearFracProvider, ICouponProvider couponProvider)
        {
            _yearFracProvider = yearFracProvider;
            _couponProvider = couponProvider;
        }

        private readonly IYearFracProvider _yearFracProvider;
        private readonly ICouponProvider _couponProvider;

        public double GetDuration(DateTime settlement, DateTime maturity, double coupon, double yield, int nFreq, DayCountBasis nBase)
        {
            double fYearfrac = _yearFracProvider.GetYearFrac(settlement, maturity, nBase);
            double fNumOfCoups = _couponProvider.GetCoupnum(settlement, maturity, nFreq, nBase);
            double fDur = 0.0;
            const double f100 = 100.0;
            coupon *= f100 / (double)nFreq;    // fCoup is used as cash flow
            yield /= nFreq;
            yield += 1.0;

            double nDiff = fYearfrac * nFreq - fNumOfCoups;

            double t;

            for (t = 1.0; t < fNumOfCoups; t++)
                fDur += (t + nDiff) * coupon / System.Math.Pow(yield, t + nDiff);

            fDur += (fNumOfCoups + nDiff) * (coupon + f100) / System.Math.Pow(yield, fNumOfCoups + nDiff);

            double p = 0.0;
            for (t = 1.0; t < fNumOfCoups; t++)
                p += coupon / System.Math.Pow(yield, t + nDiff);

            p += (coupon + f100) / System.Math.Pow(yield, fNumOfCoups + nDiff);

            fDur /= p;
            fDur /= (double)nFreq;

            return fDur;
        }
    }
}
