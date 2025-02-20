/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "7.2",
        Description = "Create custom, reusable functions and call them by a friendly name",
        IntroducedInExcelVersion = "2021")]
    internal class Lambda : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // just add the variables here
            if(arguments.Last().DataType != DataType.LambdaTokens)
            {
                return CreateResult(eErrorType.Value);
            }
            var tokens = arguments.Last().Value as List<Token>;
            var calculator = new LambdaCalculator(tokens);
            var variables = new List<CompileResult>();
            for(var i = 0; i < arguments.Count -1; i++)
            {
                var arg = arguments[i];
                var cr = new CompileResult(arg.Value, arg.DataType);
                variables.Add(cr);
            }
            calculator.SetVariables(variables);
            return new CompileResult(calculator, DataType.LambdaCalculation);
        }
    }
}
