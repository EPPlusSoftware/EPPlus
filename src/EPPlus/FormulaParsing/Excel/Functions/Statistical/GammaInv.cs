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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
    Description = "Returns the individual term binomial distribution probability.")]

    internal class GammaInv : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2);
        }
        //private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        //{
        //    ArrayParameterIndexes = new List<int> { 0, 1, 2 }
        //};
        //public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        //{
        //    return _arrayConfig;
        //}
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 3) return CompileResult.GetErrorResult(eErrorType.Value);

            var probability = ArgToDecimal(arguments, 0, out ExcelErrorValue error);
            if (error != null) return CompileResult.GetErrorResult(error.Type);

            var alpha = ArgToDecimal(arguments, 1, out error); 
            if(error != null) return CompileResult.GetErrorResult(error.Type);

            var beta = ArgToDecimal(arguments, 2, out error);
            if (error != null) return CompileResult.GetErrorResult(error.Type);

            if (probability < 0 || probability > 1 || alpha <= 0 || beta <= 0) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = GammaHelper.InverseGamma(probability, alpha) * beta;

            return CreateResult(result, DataType.Decimal);
        }
    }
}