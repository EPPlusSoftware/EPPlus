/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function (ported to c# from Microsoft.VisualBasic.Financial.vb (MIT))
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    /// <summary>
    /// Rate implementation
    /// </summary>
    public static class RateImpl
    {
        private const double cnL_IT_STEP = 0.00001;
        private const double cnL_IT_EPSILON = 0.0000001;

        /// <summary>
        /// Rate
        /// </summary>
        /// <param name="NPer"></param>
        /// <param name="Pmt"></param>
        /// <param name="PV"></param>
        /// <param name="FV"></param>
        /// <param name="Due"></param>
        /// <param name="Guess"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static FinanceCalcResult<double> Rate(double NPer, double Pmt, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod, double Guess = 0.1)
        {
            double dTemp;
            double dRate0;
            double dRate1;
            double dY0;
            double dY1;
            int I;

            // Check for error condition
            if (NPer <= 0.0)
                throw new ArgumentException("NPer must by greater than zero");

            dRate0 = Guess;
            dY0 = LEvalRate(dRate0, NPer, Pmt, PV, FV, Due);
            if (dY0 > 0)
                dRate1 = (dRate0 / 2);
            else
                dRate1 = (dRate0 * 2);

            dY1 = LEvalRate(dRate1, NPer, Pmt, PV, FV, Due);

            for (I = 0; I <= 39; I++)
            {
                if (dY1 == dY0)
                {
                    if (dRate1 > dRate0)
                        dRate0 = dRate0 - cnL_IT_STEP;
                    else
                        dRate0 = dRate0 - cnL_IT_STEP * (-1);
                    dY0 = LEvalRate(dRate0, NPer, Pmt, PV, FV, Due);
                    if (dY1 == dY0)
                        return new FinanceCalcResult<double>(eErrorType.Num);
                }

                dRate0 = dRate1 - (dRate1 - dRate0) * dY1 / (dY1 - dY0);

                // Secant method of generating next approximation
                dY0 = LEvalRate(dRate0, NPer, Pmt, PV, FV, Due);
                if (System.Math.Abs(dY0) < cnL_IT_EPSILON)
                    return new FinanceCalcResult<double>(dRate0);

                dTemp = dY0;
                dY0 = dY1;
                dY1 = dTemp;
                dTemp = dRate0;
                dRate0 = dRate1;
                dRate1 = dTemp;
            }

            return new FinanceCalcResult<double>(eErrorType.Num);
        }
        /// <summary>
        /// LEvalRate
        /// </summary>
        /// <param name="Rate"></param>
        /// <param name="NPer"></param>
        /// <param name="Pmt"></param>
        /// <param name="PV"></param>
        /// <param name="dFv"></param>
        /// <param name="Due"></param>
        /// <returns></returns>
        public static double LEvalRate(double Rate, double NPer, double Pmt, double PV, double dFv, PmtDue Due)
        {
            double dTemp1;
            double dTemp2;
            double dTemp3;

            if (Rate == 0.0)
                return (PV + Pmt * NPer + dFv);
            else
            {
                dTemp3 = Rate + 1.0;
                // WARSI Using the exponent operator for pow(..) in C code of LEvalRate. Still got
                // to make sure that they (pow and ^) are same for all conditions
                dTemp1 = System.Math.Pow(dTemp3, NPer);

                if (Due != PmtDue.EndOfPeriod)
                    dTemp2 = 1 + Rate;
                else
                    dTemp2 = 1.0;
                return (PV * dTemp1 + Pmt * dTemp2 * (dTemp1 - 1) / Rate + dFv);
            }
        }
    }
}
