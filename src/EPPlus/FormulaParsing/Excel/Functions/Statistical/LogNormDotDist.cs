﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    internal class LogNormDotDist : NormalDistributionBase
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            {
                var z = ArgToDecimal(arguments, 0);
                var cumulative = ArgToBool(arguments, 1);
                double result;
                if (cumulative)
                {
                    result = CumulativeDistribution(z, 0, 1);
                }
                else
                {
                    result = ProbabilityDensity(z, 0, 1);
                }
                return CreateResult(result, DataType.Decimal);
            }
        }
    }   
}
