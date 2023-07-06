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
    }
}
