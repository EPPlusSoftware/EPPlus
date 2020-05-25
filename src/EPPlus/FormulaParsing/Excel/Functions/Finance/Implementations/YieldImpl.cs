using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{   
    internal class YieldImpl
    {
        public YieldImpl(ICouponProvider couponProvider, IPriceProvider priceProvider)
        {
            _couponProvider = couponProvider;
            _priceProvider = priceProvider;
        }

        private readonly ICouponProvider _couponProvider;
        private readonly IPriceProvider _priceProvider;

        private static bool AreEqual(double x, double y)
        {
            return System.Math.Abs(x - y) < 0.000000001;
        }

        public double GetYield(System.DateTime settlement, System.DateTime maturity, double rate, double pr, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360)
        {
            
            var A = _couponProvider.GetCoupdaybs(settlement, maturity, frequency, basis);
            var N = _couponProvider.GetCoupnum(settlement, maturity, frequency, basis);
            var E = _couponProvider.GetCoupdays(settlement, maturity, frequency, basis);

            if (N <= -1)
            { 
                var DSR = E - A;
                var part1 = (redemption / 100 + rate / frequency) - (pr / 100d + (A / E * rate / frequency));
                var part2 = pr / 100d + (A / E) * rate / frequency;
                var retVal = part1 / part2 * ((frequency * E) / DSR);

                return retVal;
            }
            else
            {
                double price = pr;
                double fPriceN = 0.0;
                double fYield1 = 0.0;
                double fYield2 = 1.0;
                double fPrice1 = _priceProvider.GetPrice(settlement, maturity, rate, fYield1, redemption, frequency, basis);
                double fPrice2 = _priceProvider.GetPrice(settlement, maturity, rate, fYield2, redemption, frequency, basis);
                double fYieldN = (fYield2 - fYield1) * 0.5;


                for (int nIter = 0; nIter < 100 && !AreEqual(fPriceN, price); nIter++)
                {
                    fPriceN = _priceProvider.GetPrice(settlement, maturity, rate, fYieldN, redemption, frequency, basis);

                    if (AreEqual(price, fPrice1))
                        return fYield1;
                    else if (AreEqual(price, fPrice2))
                        return fYield2;
                    else if (AreEqual(price, fPriceN))
                        return fYieldN;
                    else if (price < fPrice2)
                    {
                        fYield2 *= 2.0;
                        fPrice2 = _priceProvider.GetPrice(settlement, maturity, rate, fYield2, redemption, frequency, basis);

                        fYieldN = (fYield2 - fYield1) * 0.5;
                    }
                    else
                    {
                        if (price < fPriceN)
                        {
                            fYield1 = fYieldN;
                            fPrice1 = fPriceN;
                        }
                        else
                        {
                            fYield2 = fYieldN;
                            fPrice2 = fPriceN;
                        }

                        fYieldN = fYield2 - (fYield2 - fYield1) * ((price - fPrice2) / (fPrice1 - fPrice2));
                    }
                }
                if (System.Math.Abs(price - fPriceN) > price / 100d)
                    throw new Exception("Result not precise enough");
                return fYieldN;      
            }
        }
    }
}
