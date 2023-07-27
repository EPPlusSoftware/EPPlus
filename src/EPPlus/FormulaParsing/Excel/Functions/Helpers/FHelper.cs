using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class FHelper
    {
        public static double ProbabilityDensityFunction(double x, double df1, double df2)
        {
            var arg1 = Math.Pow(df1 * x, df1) * Math.Pow(df2, df2);
            var arg2 = Math.Pow(df1 * x + df2, df1 + df2);
            var arg3 = x * BetaHelper.Beta(df1 / 2, df2 / 2);
            return Math.Sqrt(arg1 / arg2) / arg3;
        }

        public static double CumulativeDistributionFunction(double x, double df1, double df2)
        {
            return BetaHelper.IBeta(df1 * x / (df1 * x + df2), df1 / 2, df2 / 2);
        }

        public static double GetProbability(double x, double df1, double df2, bool cumulative)
        {
            var fValue = (cumulative) ? FHelper.CumulativeDistributionFunction(x, df1, df2) : FHelper.ProbabilityDensityFunction(x, df1, df2);
            return fValue;
        }
    }
}
