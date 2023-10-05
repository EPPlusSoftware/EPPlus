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
    Description = "Returns the negative binomial distribution")]


    internal class NegBinomDist : ExcelFunction
    {
        public override int ArgumentMinLength => 4;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override string NamespacePrefix => "_xlfn.";

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2, 3);
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 4) return CompileResult.GetErrorResult(eErrorType.Value);

            var numF = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            numF = Math.Floor(numF);

            var numS = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            numS = Math.Floor(numS);
            
            var probS = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            
            var cumulative = ArgToBool(arguments, 3);
            if (numF < 0 || numS < 1 || probS < 0 || probS > 1) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = 0d;
            if (cumulative)
            {
                for (var i=0; i <= numF; i++)
                {
                    var denominator = numS - 1;
                    var nominator = i + denominator;
                    var combin = MathHelper.Factorial(nominator, nominator - denominator) / MathHelper.Factorial(denominator);
                    result += combin * Math.Pow(probS, numS) * Math.Pow(1 - probS, i);
                }
            }
            else
            {
                var denominator = numS - 1;
                var nominator = numF + denominator;
                var combin = MathHelper.Factorial(nominator, nominator - denominator) / MathHelper.Factorial(denominator);
                result = combin * Math.Pow(probS, numS) * Math.Pow(1 - probS, numF);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
