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
using OfficeOpenXml.Utils.Formula;
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
            var elt = LambdaExtractor.GetLambdaTokens(_formula);
            var rpn2 = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            rpn2.ExpressionStack = _rpnFormula.ExpressionStack;
            rpn2.FunctionStack = _rpnFormula.FunctionStack;
            rpn2.IgnoreCaching = true;
            var clc = new LambdaCalculator(elt.LambdaTokens.Tokens, ctx.VariableStorage.Peek(), rpn2);
            var vs = new List<VariableCompileResult>();
            foreach (var v2 in elt.VariableTokens)
            {
                vs.Add(new VariableCompileResult(v2.Value, null, DataType.LambdaVariableDeclaration, null));
            }
            clc.SetVariables(vs, ctx);
            for (var argIx = 0; argIx < arguments.Count || argIx < clc.NumberOfVariables; argIx++)
            {
                clc.SetVariableValue(argIx, arguments[argIx].Value, arguments[argIx].DataType, ctx);
            }
            clc.BeginCalculation();
            var result = clc.Execute(ctx);
            return result;
            //var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            //formula.IgnoreCaching = true;
            //formula.ExpressionStack = _rpnFormula.ExpressionStack;
            //formula.FunctionStack = _rpnFormula.FunctionStack;
            //var tokens = SourceCodeTokenizer.Default.Tokenize(_formula).ToList();
            //var tokensRpn = FormulaExecutor.CreateRPNTokens(tokens);
            //var rpnTokens = new RpnTokens { Tokens = tokensRpn.Tokens };
            //var variables = new List<VariableCompileResult>();
            //var startIx = -1;
            //foreach(var token in rpnTokens)
            //{
            //    startIx++;
            //    if (token.TokenType == TokenType.CommaLambda) break;
            //    if(token.TokenType == TokenType.ParameterVariableDeclaration)
            //    {
            //        variables.Add(new VariableCompileResult(token.Value, null, DataType.LambdaVariableDeclaration, null));
            //    }
            //}
 
            //var calculator = new LambdaCalculator(tokensRpn.Tokens, ctx.VariableStorage.Peek(), formula);
            //calculator.SetVariables(variables, ctx);
            //calculator.BeginCalculation();
            //for(var argIx = 0; argIx < arguments.Count || argIx < calculator.NumberOfVariables; argIx++)
            //{
            //    calculator.SetVariableValue(argIx, arguments[argIx].Value, arguments[argIx].DataType, ctx);
            //}
            //calculator.IsCompileLambdaName = true;
            //return calculator.Execute(ctx);
            ////return  new CompileResult(calculator, DataType.LambdaCalculation);
            //var tokens2 = calculator.GetCurrentTokens();
            //var rpnTokens2 = new RpnTokens { Tokens = tokens2 };
            //formula.SetTokens(new RpnTokens { Tokens = tokens2 }, ctx);
            //LambdaFormulaSettings lambdaSettings = default;
            //formula._expressions = FormulaExecutor.CompileExpressions(ref lambdaSettings, ref rpnTokens2, ctx);
            //var result = RpnFormulaExecution.ExecutePartialFormula(ctx.DependencyChain, formula, ctx.CalcOption, false);
            //return result;
        }
    }
}
