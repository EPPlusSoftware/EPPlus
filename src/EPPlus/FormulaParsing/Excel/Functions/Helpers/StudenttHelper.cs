/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class StudenttHelper
    {
        public static double ProbabilityDensityFunction(double x, double degreesOfFreedom)
        {

            //PDF when the initial formula has the argument cumulative = false.

            var numeratorPDF = System.Math.Pow(degreesOfFreedom / (degreesOfFreedom + System.Math.Pow(x, 2)), (degreesOfFreedom + 1) / 2);
            var denominatorPDF = System.Math.Sqrt(degreesOfFreedom) * BetaHelper.Beta(degreesOfFreedom / 2, 0.5d);

            var pdf = numeratorPDF / denominatorPDF;
            return pdf;
        }

        public static double CumulativeDistributionFunction(double x, double degreesOfFreedom)
        {

            //Using regularized incomplete beta function to find the cumulative distribution function when initial formula has the argument cumulative = true.

            var cdf = 0d;

            if (x <= 0)
            {
                var arg1 = degreesOfFreedom / (System.Math.Pow(x, 2) + degreesOfFreedom);
                var arg2 = degreesOfFreedom / 2;
                var arg3 = 0.5d;
                cdf = 0.5d * BetaHelper.IBeta(arg1, arg2, arg3);
            }
            else
            {
                var arg1 = System.Math.Pow(x, 2) / (System.Math.Pow(x, 2) + degreesOfFreedom);
                var arg2 = 0.5d;
                var arg3 = degreesOfFreedom / 2;
                cdf = 0.5d * (BetaHelper.IBeta(arg1, arg2, arg3) + 1);
            }

            return cdf;
        }

        public static double InverseTFunc(double probability, double degreesOfFreedom)
        {
            //Approximating the inverse integral with bisection algorithm. Methods like newton-rhapson might be
            //better, but this is a rather nice looking and easy implementation.
            //When the cdf gets close enough to the probability, we know that we have found the area that correlates to given probability.

            var epsilon = 0.000000000001;
            var lBound = -500d;
            var uBound = 500d;

            while (uBound - lBound > epsilon)
            {
                var intersect = (lBound + uBound) / 2d;
                var cdf = CumulativeDistributionFunction(intersect, degreesOfFreedom);

                var diff = Math.Abs(cdf - probability);
                if (Math.Abs(cdf - probability) < epsilon)
                {
                    return intersect;
                }
                else if (cdf < probability)
                {
                    lBound = intersect;
                }
                else
                {
                    uBound = intersect;
                }

            }

            return (lBound + uBound) / 2;
        }
    }
}
