/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    /// <summary>
    /// Xirr implementation
    /// </summary>
    public static class XirrImpl
    {
        /// <summary>
        /// Get Xirr
        /// </summary>
        /// <param name="aValues"></param>
        /// <param name="aDates"></param>
        /// <param name="rGuessRate"></param>
        /// <returns></returns>
        public static FinanceCalcResult<double> GetXirr(IEnumerable<double> aValues, IEnumerable<DateTime> aDates, double rGuessRate = 0.1)
        {
            if (aValues.Count() != aDates.Count()) return new FinanceCalcResult<double>(eErrorType.Value);

            // maximum epsilon for end of iteration
            const double fMaxEps = 1e-10;
            // maximum number of iterations
            const int nMaxIter = 50;

            // Newton's method - try to find a fResultRate, so that lcl_sca_XirrResult() returns 0.
            int nIter = 0;
            double fResultRate = rGuessRate;
            double fResultValue;
            int nIterScan = 0;
            bool bContLoop = false;
            bool bResultRateScanEnd = false;

            // First the inner while-loop will be executed using the default Value fResultRate
            // or the user guessed fResultRate if those do not deliver a solution for the
            // Newton's method then the range from -0.99 to +0.99 will be scanned with a
            // step size of 0.01 to find fResultRate's value which can deliver a solution
            do
            {
                if (nIterScan >= 1)
                    fResultRate = -0.99 + (nIterScan - 1) * 0.01;
                do
                {
                    fResultValue = lcl_sca_XirrResult(aValues, aDates, fResultRate);
                    double fNewRate = fResultRate - fResultValue / lcl_sca_XirrResult_Deriv1(aValues, aDates, fResultRate);
                    double fRateEps = System.Math.Abs(fNewRate - fResultRate);
                    fResultRate = fNewRate;
                    bContLoop = (fRateEps > fMaxEps) && (System.Math.Abs(fResultValue) > fMaxEps);
                }
                while (bContLoop && (++nIter < nMaxIter));
                nIter = 0;
                if (double.IsNaN(fResultRate) || double.IsInfinity(fResultRate)
                    || double.IsNaN(fResultValue) || double.IsInfinity(fResultValue))
                    bContLoop = true;

                ++nIterScan;
                bResultRateScanEnd = (nIterScan >= 200);
            }
            while (bContLoop && !bResultRateScanEnd);

            if (bContLoop)
                return new FinanceCalcResult<double>(OfficeOpenXml.eErrorType.Value);
            return new FinanceCalcResult<double>(fResultRate);
        }

        static double lcl_sca_XirrResult_Deriv1(IEnumerable<double> rValues, IEnumerable<DateTime> rDates, double fRate)
        {
            /*  V_0 ... V_n = input values.
                D_0 ... D_n = input dates.
                R           = input interest rate.
      
                r   := R+1
                E_i := (D_i-D_0) / 365
      
                                    n    V_i
                f'(R)  =  [ V_0 + SUM   ------- ]'
                                    i=1  r^E_i
      
                                n           V_i                 n    E_i V_i
                        =  0 + SUM   -E_i ----------- r'  =  - SUM   ----------- .
                                i=1       r^(E_i+1)             i=1  r^(E_i+1)
            */
            var D_0 = rDates.ElementAt(0);
            double r = fRate + 1.0;
            double fResult = 0.0;
            for (int i = 1, nCount = rValues.Count(); i < nCount; ++i)
            {
                var E_i = (rDates.ElementAt(i).Subtract(D_0).TotalDays / 365d);
                fResult -= E_i * rValues.ElementAt(i) / System.Math.Pow(r, E_i + 1.0);
            }
            return fResult;
        }

        static double lcl_sca_XirrResult(IEnumerable<double> rValues, IEnumerable<DateTime> rDates, double fRate)
        {
            /*  V_0 ... V_n = input values.
                D_0 ... D_n = input dates.
                R           = input interest rate.
  
                r   := R+1
                E_i := (D_i-D_0) / 365
  
                            n    V_i                n    V_i
                f(R)  =  SUM   -------  =  V_0 + SUM   ------- .
                            i=0  r^E_i              i=1  r^E_i
            */
            var D_0 = rDates.ElementAt(0);
            double r = fRate + 1.0;
            double fResult = rValues.ElementAt(0);
            for (int i = 1, nCount = rValues.Count(); i < nCount; ++i)
                fResult += rValues.ElementAt(i) / System.Math.Pow(r, ((rDates.ElementAt(i).Subtract(D_0).TotalDays) / 365.0));
            return fResult;
        }
    }
}
