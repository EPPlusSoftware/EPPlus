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
using System;

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
