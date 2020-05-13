/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class IrrImpl
    {
        private const double cnL_IT_STEP = 0.00001;
        private const double cnL_IT_EPSILON = 0.0000001;

        private static double OptPV2(ref double[] ValueArray, double Guess = 0.1)
        {
            int lUpper, lLower, lIndex;

            lLower = 0;
            lUpper = ValueArray.Length - 1;

            double dTotal = 0.0;
            double divRate = 1.0 + Guess;

            while ((lLower <= lUpper) && ValueArray[lLower] == 0.0)
                lLower = lLower + 1;
            for (lIndex = lUpper; lIndex > lLower - 1; lIndex--)
            {
                dTotal = dTotal / divRate;
                dTotal = dTotal + ValueArray[lIndex];
            }
            return dTotal;
        }

        internal static FinanceCalcResult Irr(double[] ValueArray, double Guess = 0.1)
        {
            double dTemp;
            double dRate0;
            double dRate1;
            double dNPv0;
            double dNpv1;
            double dNpvEpsilon;
            double dTemp1;
            int lIndex;
            int lCVal;
            int lUpper = 0;

            // Compiler assures that rank of ValueArray is always 1, no need to check it.  
            // WARSI Check for error codes returned by UBound. Check if they match with C code
            try
            {
                //Needed to catch dynamic arrays which have not been constructed yet
                lUpper = ValueArray.Length - 1;
            }
            catch (StackOverflowException soe)
            {
                return new FinanceCalcResult(eErrorType.Value);
            }
            catch (OutOfMemoryException ome)
            {
                return new FinanceCalcResult(eErrorType.Value);
            }
            catch (ArgumentException ae)
            {
                // return error due to invalid value array
                return new FinanceCalcResult(eErrorType.Value);
            }

            lCVal = lUpper + 1;

            //Function fails for invalid parameters
            if (Guess <= -1.0)
            {
                return new FinanceCalcResult(eErrorType.Num);
            }
            if (lCVal <= 1)
            {
                return new FinanceCalcResult(eErrorType.Num);
            }

            //'We scale the epsilon depending on cash flow values. It is necessary
            //'because even in max accuracy case where diff is in 16th digit it
            //'would get scaled up.
            if (ValueArray[0] > 0.0)
            {
                dTemp = ValueArray[0];
            }
            else
            {
                dTemp = -ValueArray[0];
            }

            for (lIndex = 0; lIndex <= lUpper; lIndex++)
            {
                //Get max of values in cash flow
                if (ValueArray[lIndex] > dTemp)
                    dTemp = ValueArray[lIndex];
                else if (-ValueArray[lIndex] > dTemp)
                    dTemp = -ValueArray[lIndex];
            }

            dNpvEpsilon = dTemp * cnL_IT_EPSILON * 0.01;

            // Set up the initial values for the secant method
            dRate0 = Guess;
            dNPv0 = OptPV2(ref ValueArray, dRate0);

            if (dNPv0 > 0.0)
                dRate1 = dRate0 + cnL_IT_STEP;
            else
                dRate1 = dRate0 - cnL_IT_STEP;

            if (dRate1 <= -1.0)
                return new FinanceCalcResult(eErrorType.Num);

            dNpv1 = OptPV2(ref ValueArray, dRate1);

            for (lIndex = 0; lIndex <= 39; lIndex++)
            {
                if (dNpv1 == dNPv0)
                {
                    if (dRate1 > dRate0)
                        dRate0 = dRate0 - cnL_IT_STEP;
                    else
                        dRate0 = dRate0 + cnL_IT_STEP;
                }
                dNPv0 = OptPV2(ref ValueArray, dRate0);
                if (dNpv1 == dNPv0)
                    return new FinanceCalcResult(eErrorType.Value);

                dRate0 = dRate1 - (dRate1 - dRate0) * dNpv1 / (dNpv1 - dNPv0);

                //Secant method of generating next approximation
                if (dRate0 <= -1.0)
                    dRate0 = (dRate1 - 1.0) * 0.5;

                //Basically give the algorithm a second chance. Helps the
                //algorithm when it starts to diverge to -ve side
                dNPv0 = OptPV2(ref ValueArray, dRate0);
                if (dRate0 > dRate1)
                    dTemp = dRate0 - dRate1;
                else
                    dTemp = dRate1 - dRate0;

                if (dNPv0 > 0.0)
                    dTemp1 = dNPv0;
                else
                    dTemp1 = -dNPv0;

                //Test : npv - > 0 and rate converges
                if (dTemp1 < dNpvEpsilon && dTemp < cnL_IT_EPSILON)
                    return new FinanceCalcResult(dRate0);

                //Exchange the values - store the new values in the 1's
                dTemp = dNPv0;
                dNPv0 = dNpv1;
                dNpv1 = dTemp;
                dTemp = dRate0;
                dRate0 = dRate1;
                dRate1 = dTemp;
            }
            return new FinanceCalcResult(eErrorType.Value);
        }
    }
}
