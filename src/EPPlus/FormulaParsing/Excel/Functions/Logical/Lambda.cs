/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2024         EPPlus Software AB       Initial release EPPlus 8.0
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "8.1",
        Description = "Create custom, reusable functions and call them by a friendly name",
        IntroducedInExcelVersion = "2021")]
    internal class Lambda : ExcelFunction
    {
        private RpnFormula _rpnFormula;
        public override int ArgumentMinLength => 2;
        public override bool IsVolatile => true;

        public override string NamespacePrefix => "_xlfn.";

        internal override void SetRpnFormula(RpnFormula formula)
        {
           _rpnFormula = formula;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // just add the variables here
            if(arguments.Last().DataType != DataType.LambdaTokens)
            {
                return CreateResult(eErrorType.Value);
            }
            var tokenResult = arguments.Last().Value as LambdaTokensResult;
            var calculator = new LambdaCalculator(tokenResult.Tokens, tokenResult.Scope, _rpnFormula);
            var variables = new List<VariableCompileResult>();
            for(var i = 0; i < arguments.Count -1; i++)
            {
                var arg = arguments[i];
                if(arg.IsVariableResult)
                {
                    variables.Add(arg.ValueAsVariableCompileResult);
                }
                else if(arg.DataType == DataType.LambdaVariableDeclaration)
                {
                    variables.Add(new VariableCompileResult(arg.Value.ToString(), null, DataType.LambdaVariableDeclaration, null));
                }
            }
            calculator.SetVariables(variables, context);
            return new CompileResult(calculator, DataType.LambdaCalculation);
        }
    }
}
