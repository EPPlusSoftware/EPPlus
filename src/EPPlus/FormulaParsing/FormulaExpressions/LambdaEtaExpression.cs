/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/20/2025         EPPlus Software AB       Initial release EPPlus 8.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaEtaExpression : VariableFunctionExpression
    {
        public LambdaEtaExpression(string tokenValue, ParsingContext ctx, int pos)
            : base(tokenValue, ctx,pos)
        {
            _tokenValue = tokenValue;
            _scope = Context.VariableStorage.GetById(VariableScopeId);
        }

        private readonly string _tokenValue;
        private readonly VariableStorageScope _scope;
        private RpnFormula _rpnFormula;

        internal override ExpressionType ExpressionType => ExpressionType.LambdaCalculation;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

        internal override void SetRpnFormula(RpnFormula formula)
        {
            _rpnFormula = formula;
        }

        public override CompileResult Compile()
        {
            if(string.IsNullOrEmpty(_tokenValue) || !_tokenValue.StartsWith("_xleta."))
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var functionName = _tokenValue.Replace("_xleta.", string.Empty);
            var func = Context.Configuration.FunctionRepository.GetFunction(functionName);
            if(func == null || func.ArgumentMinLength > 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var paramName = $"p{Guid.NewGuid().ToString("N")}";
            var formula = $"LAMBDA({paramName}, {functionName}({paramName}))";
            var lambdaTokens = SourceCodeTokenizer.Default.Tokenize(formula);
            var rpnTokens = FormulaExecutor.CreateRPNTokens(lambdaTokens);

            var rpnFormula = new RpnFormula(Context.CurrentWorksheet, Context.CurrentCell.Row, Context.CurrentCell.Column, Context.VariableStorage)
            {
                _tokens = rpnTokens,
                ExpressionStack = _rpnFormula.ExpressionStack,
                FunctionStack = _rpnFormula.FunctionStack
            };
            rpnFormula.SetTokens(rpnTokens, Context, _scope);
            var result = RpnFormulaExecution.ExecutePartialFormula(Context.DependencyChain, rpnFormula, Context.CalcOption, false);
            if(result.DataType != DataType.LambdaCalculation)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            if (result.ResultValue is LambdaCalculator calculator)
            {
                calculator.SetEtaReducedLambdaFunction(functionName, Context);
            }
            return result;
        }

        public override Expression Negate()
        {
            throw new NotImplementedException();
        }
    }
}
