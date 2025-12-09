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
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class ChiSquareHelper
    {
        public static double PropbabilityDistribution(double n, double degreeOfFreedom)
        {
            if (n < 0d)
            {
                return 0d;
            }
            else if(n == 0d && degreeOfFreedom == 2d)
            {
                return 0.5d;
            }
            return System.Math.Exp((degreeOfFreedom / 2 - 1) * System.Math.Log(n) - n / 2 - (degreeOfFreedom / 2d) * System.Math.Log(2d) - GammaHelper.LogGamma(degreeOfFreedom / 2));

        }

        public static double CumulativeDistribution(double n, double degreeOfFreedom)
        {
            if(n < 0d)
            {
                return 0;
            }
            return GammaHelper.RegularizedGammaP(degreeOfFreedom / 2, n / 2, 1.0e-15, 10000);
        }

        public static double Inverse(double n, double degreeOfFreedom)
        {
            return 2 * GammaPinvHelper.gammapinv(n, 0.5 * degreeOfFreedom);
        }
    }
}
