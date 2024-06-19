/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations
{
    /// <summary>
    /// Bessel K
    /// </summary>
    internal class BesselKImpl : BesselBase
    {
        static FinanceCalcResult<double> Besselk0(double fNum)
        {
            double fRet;

            if (fNum <= 2.0)
            {
                double fNum2 = fNum * 0.5;
                double y = fNum2 * fNum2;

                var iResult = new BesselIimpl().BesselI(fNum, 0);
                if (iResult.HasError) return iResult;

                fRet = -System.Math.Log(fNum2) * iResult.Result +
                        (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.3488590e-1 +
                            y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))));
            }
            else
            {
                double y = 2.0 / fNum;

                fRet = System.Math.Exp(-fNum) / System.Math.Sqrt(fNum) * (1.25331414 + y * (-0.7832358e-1 +
                        y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 +
                        y * (-0.251540e-2 + y * 0.53208e-3))))));
            }

            return new FinanceCalcResult<double>(fRet);
        }

        /// @throws IllegalArgumentException
        /// @throws NoConvergenceException
        static FinanceCalcResult<double> Besselk1(double fNum)
        {
            double fRet;

            if (fNum <= 2.0)
            {
                double fNum2 = fNum * 0.5;
                double y = fNum2 * fNum2;

                var iResult = new BesselIimpl().BesselI(fNum, 1);
                if (iResult.HasError) return iResult;
                fRet = System.Math.Log(fNum2) * iResult.Result +
                        (1.0 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 +
                            y * (-0.110404e-2 + y * -0.4686e-4))))))
                        / fNum;
            }
            else
            {
                double y = 2.0 / fNum;

                fRet = System.Math.Exp(-fNum) / System.Math.Sqrt(fNum) * (1.25331414 + y * (0.23498619 +
                        y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 +
                        y * (0.325614e-2 + y * -0.68245e-3))))));
            }

            return new FinanceCalcResult<double>(fRet);
        }
        /// <summary>
        /// Bessel K
        /// </summary>
        /// <param name="fNum"></param>
        /// <param name="nOrder"></param>
        /// <returns></returns>
        public FinanceCalcResult<double> BesselK(double fNum, int nOrder)
        {
            switch (nOrder)
            {
                case 0: return Besselk0(fNum);
                case 1: return Besselk1(fNum);
                default:
                    {
                        var k0Result = Besselk0(fNum);
                        if (k0Result.HasError) return k0Result;
                        var k1Result = Besselk1(fNum);
                        if (k1Result.HasError) return k1Result;
                        double fTox = 2.0 / fNum;
                        double fBkm = k0Result.Result;
                        double fBk = k1Result.Result;

                        for (int n = 1; n < nOrder; n++)
                        {
                            var fBkp = fBkm + n * fTox * fBk;
                            fBkm = fBk;
                            fBk = fBkp;
                        }

                        return new FinanceCalcResult<double>(fBk);
                    }
            }
        }

    }
}
