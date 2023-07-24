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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
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
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
       
        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1, 2, 3 }
        };
        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            {
                if (arguments.Count > 4)
                {
                    return CompileResult.GetErrorResult(eErrorType.Value);
                }
                var z = ArgToDecimal(arguments, 0);
                var mean = ArgToDecimal(arguments, 1);
                var stdev = ArgToDecimal(arguments, 2);
                var cumulative = ArgToBool(arguments, 3);
                if (stdev <= 0)
                {
                    return CompileResult.GetErrorResult(eErrorType.Num);
                }

                var result = 0d;
                if (cumulative)
                {
                    result = CumulativeDistribution(Math.Log(z), mean, stdev);
                }
                else
                {
                    result = ProbabilityDensity(Math.Log(z), mean, stdev)/z;
                }
                return CreateResult(result, DataType.Decimal);
            }
        }
    }   
}
