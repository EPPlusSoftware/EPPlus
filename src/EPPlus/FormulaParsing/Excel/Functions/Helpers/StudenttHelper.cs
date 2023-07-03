using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class StudenttHelper
    {
        public static double PDF(double x, double degreesOfFreedom)
        {
            var term1 = GammaHelper.gamma((degreesOfFreedom + 1) / 2);
            var term2 = System.Math.Sqrt(degreesOfFreedom * System.Math.PI) * GammaHelper.gamma(degreesOfFreedom / 2);
            var term3 = System.Math.Pow(1 + System.Math.Pow(x, 2) / degreesOfFreedom, -1 * (degreesOfFreedom + 1) / 2);

            var probabilityDensityFunction = term1 / term2 * term3;
            return probabilityDensityFunction;
        }

        public static double CDF(double x, double degreesOfFreedom)
        {
            //Cumulative dist function is cumulative pdf

            var cumulativeDistributionFunction = 0d;

            for (var i = 0; i <= System.Math.Floor(x);  i++)
            {
                cumulativeDistributionFunction += PDF(i / 1000, degreesOfFreedom);
            }

            return cumulativeDistributionFunction;
        }
    }
}
