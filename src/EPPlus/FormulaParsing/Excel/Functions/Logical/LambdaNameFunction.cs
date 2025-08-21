/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/03/2025         EPPlus Software AB       EPPlus 8.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.DependencyChain;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    internal class LambdaNameFunction : ExcelFunction
    {
        public LambdaNameFunction(string formula)
        {
            _formula = formula;
        }

        private readonly string _formula;
        private RpnFormula _rpnFormula;

        public override int ArgumentMinLength => 0;

        public override bool IsVolatile => true;

        internal override void SetRpnFormula(RpnFormula formula)
        {
            _rpnFormula = formula;
        }

        //public override CompileResult Execute2(IList<FunctionArgument> arguments, ParsingContext context)
        //{
        //    var tokens = SourceCodeTokenizer.Default.Tokenize(_formula).ToList();
        //    var tokensRpn = FormulaExecutor.CreateRPNTokens(tokens);
        //    var calculator = new LambdaCalculator(tokensRpn.Tokens, context.VariableStorage.Peek());
        //    var variables = new List<VariableCompileResult>();
        //    for (var i = 0; i < arguments.Count; i++)
        //    {
        //        var arg = arguments[i];
        //        if (arg.IsVariableResult)
        //        {
        //            variables.Add(arg.ValueAsVariableCompileResult);
        //        }
        //        else if (arg.DataType == DataType.LambdaVariableDeclaration)
        //        {
        //            variables.Add(new VariableCompileResult(arg.Value.ToString(), null, DataType.LambdaVariableDeclaration, null));
        //        }
        //    }
        //    calculator.SetVariables(variables, context);
        //    return new CompileResult(calculator, DataType.LambdaCalculation);
        //}

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext ctx)
        {
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            formula.IgnoreCaching = true;
            formula.ExpressionStack = _rpnFormula.ExpressionStack;
            formula.FunctionStack = _rpnFormula.FunctionStack;
            var tokens = SourceCodeTokenizer.Default.Tokenize(_formula).ToList();
            var tokensRpn = FormulaExecutor.CreateRPNTokens(tokens);
            var rpnTokens = new RpnTokens { Tokens = tokensRpn.Tokens };
            var variables = new List<VariableCompileResult>();
            foreach(var token in rpnTokens)
            {
                if (token.TokenType == TokenType.CommaLambda) break;
                if(token.TokenType == TokenType.ParameterVariableDeclaration)
                {
                    variables.Add(new VariableCompileResult(token.Value, null, DataType.LambdaVariableDeclaration, null));
                }
            }
            //var chain = new RpnOptimizedDependencyChain(ctx.CurrentWorksheet.Workbook, ctx.CalcOption);
            //formula.SetFormula(_formula, chain);
            //var cr = RpnFormulaExecution.ExecutePartialFormula(chain, formula, ctx.CalcOption, false);
            //if (cr.DataType != DataType.LambdaCalculation) return CompileResult.GetErrorResult(eErrorType.Value);
            //var calculator = cr.Result as LambdaCalculator;
            var calculator = new LambdaCalculator(tokensRpn.Tokens, ctx.VariableStorage.Peek(), formula);
            calculator.SetVariables(variables, ctx);
            calculator.BeginCalculation();
            for(var argIx = 0; argIx < arguments.Count || argIx < calculator.NumberOfVariables; argIx++)
            {
                calculator.SetVariableValue(argIx, arguments[argIx].Value, arguments[argIx].DataType, ctx);
            }
            calculator.IsCompileLambdaName = true;
            //return  new CompileResult(calculator, DataType.LambdaCalculation);
            var tokens2 = calculator.GetCurrentTokens();
            var rpnTokens2 = new RpnTokens { Tokens = tokens2 };
            formula.SetTokens(new RpnTokens { Tokens = tokens2 }, ctx);
            LambdaFormulaSettings lambdaSettings = default;
            formula._expressions = FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens2, ctx);
            var result = RpnFormulaExecution.ExecutePartialFormula(ctx.DependencyChain, formula, ctx.CalcOption, false);
            ctx.VariableStorage.Pop();
            return result;
        }
    }
}
