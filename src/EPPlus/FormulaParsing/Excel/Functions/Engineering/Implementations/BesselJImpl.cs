/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations
{
    /// <summary>
    /// Bessel J
    /// </summary>
    internal class BesselJImpl : BesselBase
    {
        /// <summary>
        /// Bessel J
        /// </summary>
        /// <param name="x"></param>
        /// <param name="N"></param>
        /// <returns></returns>
        public FinanceCalcResult<double> BesselJ(double x, int N)
        {
            if (N < 0)
                return new FinanceCalcResult<double>(eErrorType.Num);
            if (x == 0.0)
                return new FinanceCalcResult<double>((N == 0) ? 1.0 : 0.0);

            /*  The algorithm works only for x>0, therefore remember sign. BesselJ
                with integer order N is an even function for even N (means J(-x)=J(x))
                and an odd function for odd N (means J(-x)=-J(x)).*/
            double fSign = (N % 2 == 1 && x < 0) ? -1.0 : 1.0;
            double fX = System.Math.Abs(x);

            const double fMaxIteration = 9000000.0; //experimental, for to return in < 3 seconds
            double fEstimateIteration = fX * 1.5 + N;
            bool bAsymptoticPossible = System.Math.Pow(fX, 0.4) > N;
            if (fEstimateIteration > fMaxIteration)
            {
                if (!bAsymptoticPossible)
                    return new FinanceCalcResult<double>(eErrorType.Num);
                var res = fSign * System.Math.Sqrt(f_2_DIV_PI / fX) * System.Math.Cos(fX - N * f_PI_DIV_2 - f_PI_DIV_4);
                return new FinanceCalcResult<double>(res);
            }

            const double epsilon = 1.0e-15; // relative error
            bool bHasfound = false;
            double k = 0.0;
            // e_{-1} = 0; e_0 = alpha_0 / b_2
            double u; // u_0 = e_0/f_0 = alpha_0/m_0 = alpha_0

            // first used with k=1
            double m_bar;         // m_bar_k = m_k * f_bar_{k-1}
            double g_bar;         // g_bar_k = m_bar_k - a_{k+1} + g_{k-1}
            double g_bar_delta_u; // g_bar_delta_u_k = f_bar_{k-1} * alpha_k
                                  // - g_{k-1} * delta_u_{k-1} - m_bar_k * u_{k-1}
                                  // f_{-1} = 0.0; f_0 = m_0 / b_2 = 1/(-1) = -1
            double g = 0.0;       // g_0= f_{-1} / f_0 = 0/(-1) = 0
            double delta_u = 0.0; // dummy initialize, first used with * 0
            double f_bar = -1.0;  // f_bar_k = 1/f_k, but only used for k=0

            if (N == 0)
            {
                //k=0; alpha_0 = 1.0
                u = 1.0; // u_0 = alpha_0
                         // k = 1.0; at least one step is necessary
                         // m_bar_k = m_k * f_bar_{k-1} ==> m_bar_1 = 0.0
                g_bar_delta_u = 0.0;    // alpha_k = 0.0, m_bar = 0.0; g= 0.0
                g_bar = -2.0 / fX;       // k = 1.0, g = 0.0
                delta_u = g_bar_delta_u / g_bar;
                u = u + delta_u;       // u_k = u_{k-1} + delta_u_k
                g = -1.0 / g_bar;       // g_k=b_{k+2}/g_bar_k
                f_bar = f_bar * g;      // f_bar_k = f_bar_{k-1}* g_k
                k = 2.0;
                // From now on all alpha_k = 0.0 and k > N+1
            }
            else
            {   // N >= 1 and alpha_k = 0.0 for k<N
                u = 0.0; // u_0 = alpha_0
                for (k = 1.0; k <= N - 1; k = k + 1.0)
                {
                    m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
                    g_bar_delta_u = -g * delta_u - m_bar * u; // alpha_k = 0.0
                    g_bar = m_bar - 2.0 * k / fX + g;
                    delta_u = g_bar_delta_u / g_bar;
                    u = u + delta_u;
                    g = -1.0 / g_bar;
                    f_bar = f_bar * g;
                }
                // Step alpha_N = 1.0
                m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
                g_bar_delta_u = f_bar - g * delta_u - m_bar * u; // alpha_k = 1.0
                g_bar = m_bar - 2.0 * k / fX + g;
                delta_u = g_bar_delta_u / g_bar;
                u = u + delta_u;
                g = -1.0 / g_bar;
                f_bar = f_bar * g;
                k = k + 1.0;
            }

            do
            {
                m_bar = 2.0 * ((k - 1.0) % 2.0) * f_bar;
                g_bar_delta_u = -g * delta_u - m_bar * u;
                g_bar = m_bar - 2.0 * k / fX + g;
                delta_u = g_bar_delta_u / g_bar;
                u = u + delta_u;
                g = -1.0 / g_bar;
                f_bar = f_bar * g;
                bHasfound = (System.Math.Abs(delta_u) <= System.Math.Abs(u) * epsilon);
                k = k + 1.0;
            }
            while (!bHasfound && k <= fMaxIteration);
            if (!bHasfound)
                return new FinanceCalcResult<double>(eErrorType.Num); // unlikely to happen

            return new FinanceCalcResult<double>(u * fSign);
        }
    }
}
