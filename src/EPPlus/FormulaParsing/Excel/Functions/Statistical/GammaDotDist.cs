/*************************************************************************************************
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
    [FunctionMetadata(
        SupportsArrays = true,
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7.0",
        Description = "Returns the gamma distribution.")]


    internal class GammaDotDist : ExcelFunction
    {
        public override int ArgumentMinLength => 4;

        public override string NamespacePrefix => "_xlfn.";
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
            if (arguments.Count > 4)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var z = ArgToDecimal(arguments, 0);
            var alpha = ArgToDecimal(arguments, 1);
            var beta = ArgToDecimal(arguments, 2);
            var cumulative = ArgToBool(arguments, 3);
            if (z < 0 || alpha<=0 || beta<=0)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var result = 0d;
            if (cumulative)
            {
                result = GammaHelper.LowerRegularizedIncompleteGamma(alpha, z/beta);
            }
            else
            {
                result = (1/(Math.Pow(beta, alpha)* GammaHelper.gamma(alpha))) * Math.Pow(z, alpha - 1) * Math.Pow(Math.E,-z/beta);
            }
            
            return CreateResult(result, DataType.Decimal);
        }
    }
}