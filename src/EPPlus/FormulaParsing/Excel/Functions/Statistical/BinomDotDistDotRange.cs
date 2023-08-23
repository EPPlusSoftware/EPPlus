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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        SupportsArrays = true,
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7.0",
        Description = "Returns the probability of a trial result using a binomial distribution.")]


    internal class BinomDotDistDotRange : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override string NamespacePrefix => "_xlfn.";
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1, 2}
        };
        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 4) return CompileResult.GetErrorResult(eErrorType.Value); 

            var trails = ArgToDecimal(arguments, 0);
            trails = Math.Floor(trails);
            var probS = ArgToDecimal(arguments, 1);
            var numS = ArgToDecimal(arguments, 2);
            numS = Math.Floor(numS);
            
            if (trails < 0 || probS < 0 || probS > 1|| numS<0|| numS>trails) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = 0d;
            if (arguments.Count > 3)
            {
                var numS2 = ArgToDecimal(arguments, 3);
                numS2 = Math.Floor(numS2);
                if (numS2 < numS || numS2 > trails) return CompileResult.GetErrorResult(eErrorType.Num);

                for (int i = (int)numS; i <= numS2; i++)
                {
                    var combin = MathHelper.Factorial(trails, trails - i) / MathHelper.Factorial(i);
                    result += combin * Math.Pow(probS, i) * Math.Pow(1 - probS, trails - i);
                }
            }
            else
            {
                var combin = MathHelper.Factorial(trails, trails - numS) / MathHelper.Factorial(numS);
                result = combin * Math.Pow(probS, numS) * Math.Pow(1 - probS, trails - numS);
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
