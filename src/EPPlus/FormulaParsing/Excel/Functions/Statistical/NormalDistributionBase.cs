/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/29/2021         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    internal abstract class NormalDistributionBase : ExcelFunction
    {
        protected double CumulativeDistribution(double x, double mean, double stdDev)
        {
            return 0.5 * (1 + ErfHelper.Erf((x - mean) / System.Math.Sqrt(2 * System.Math.Pow(stdDev, 2))));
        }

        protected double ProbabilityDensity(double x, double mean, double stdDev)
        {
            return System.Math.Exp(-0.5 * System.Math.Log(2 * System.Math.PI) - System.Math.Log(stdDev) - System.Math.Pow(x - mean, 2) / (2 * System.Math.Pow(stdDev, 2)));
        }
    }
}
