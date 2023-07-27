/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
            SupportsArrays = true,
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2010",
            Description = "Calculates the inverse of the Cumulative Normal Distribution Function for a supplied value of x, and a supplied distribution mean & standard deviation. Note that this is the same implementation as NORMINV.")]
    internal class NormDotSdotDist : NormalDistributionBase
    {
        public override int ArgumentMinLength => 2;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1 }
        };
        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
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
