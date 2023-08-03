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
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the individual term binomial distribution probability.")]


    internal class BinomDist : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            {
                var numberS = ArgToDecimal(arguments, 0);
                numberS = Math.Floor(numberS);
                var trails = ArgToDecimal(arguments, 1);
                trails = Math.Floor(trails);
                var probabilityS = ArgToDecimal(arguments, 2);
                var cumulative = ArgToBool(arguments, 3);

                if (arguments.Count >4) return CompileResult.GetErrorResult(eErrorType.Value);
                if (numberS < 0 || numberS>trails || probabilityS<0 || probabilityS>1) return CompileResult.GetErrorResult(eErrorType.Num);

                var result = 0d;
                if (cumulative)
                {
                    for (var i = 0; i <= numberS; i++)
                    {
                        var combin = MathHelper.Factorial(trails, trails - i) / MathHelper.Factorial(i);
                        result += combin*Math.Pow(probabilityS, i) * Math.Pow(1 - probabilityS, trails - i);
                    }
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
}