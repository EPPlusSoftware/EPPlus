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

    internal class BinomDist : ExcelFunction
    {
        public override int ArgumentMinLength => 4;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2, 3);
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 4) return CompileResult.GetErrorResult(eErrorType.Value);

            var numberS = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            numberS = Math.Floor(numberS);

            var trails = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            trails = Math.Floor(trails);

            var probabilityS = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var cumulative = ArgToBool(arguments, 3);

            if (numberS < 0 || numberS > trails || probabilityS < 0 || probabilityS > 1) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = 0d;
            if (cumulative)
            {
                result = BinomHelper.CumulativeDistrubution(numberS, trails, probabilityS);
            }
            else
            {
                var combin = MathHelper.Factorial(trails, trails - numberS) / MathHelper.Factorial(numberS);
                result = combin * Math.Pow(probabilityS, numberS) * Math.Pow(1 - probabilityS, trails - numberS);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}