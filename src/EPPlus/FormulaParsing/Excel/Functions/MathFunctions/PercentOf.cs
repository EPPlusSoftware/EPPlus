/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/11/2025         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
          Category = ExcelFunctionCategory.MathAndTrig,
          EPPlusVersion = "8.0",
          Description = "The PERCENTOF function sums the values in the subset and divides it by all the values.",
          SupportsArrays = true)]
    internal class PercentOf : ExcelFunction
    {       
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 2)
                return CompileResult.GetErrorResult(eErrorType.Value); 

            var v1 = SumV2.Calculate(arguments[0], context, out eErrorType? e1);
            var v2 = SumV2.Calculate(arguments[1], context, out eErrorType? e2);

            if (v2 == 0) return CompileResult.GetErrorResult(eErrorType.Div0);                
            if (e1 != null) return CompileResult.GetErrorResult(eErrorType.Num);
            if (e2 != null) return CompileResult.GetErrorResult(eErrorType.Num);

            var result = v1 / v2;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
