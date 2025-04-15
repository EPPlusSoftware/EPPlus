using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    internal class LambdaNameFunction : ExcelFunction
    {
        public LambdaNameFunction(string formula)
        {
            _formula = formula;
        }

        private readonly string _formula;

        public override int ArgumentMinLength => 0;

        public override bool IsVolatile => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext ctx)
        {
            var formula = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column);
            var tokens = SourceCodeTokenizer.Default.Tokenize(_formula).ToList();
            var rpnTokens = new RpnTokens { Tokens = tokens };
            var chain = new RpnOptimizedDependencyChain(ctx.CurrentWorksheet.Workbook, ctx.CalcOption);
            formula.SetFormula(_formula, chain);
            var cr = RpnFormulaExecution.ExecutePartialFormula(chain, formula, ctx.CalcOption, false);
            if (cr.DataType != DataType.LambdaCalculation) return CompileResult.GetErrorResult(eErrorType.Value);
            var calculator = cr.Result as LambdaCalculator;
            calculator.BeginCalculation();
            for(var argIx = 0; argIx < arguments.Count || argIx < calculator.NumberOfVariables; argIx++)
            {
                calculator.SetVariableValue(argIx, arguments[argIx].Value, arguments[argIx].DataType, ctx);
            }
            var r2 = calculator.Execute(ctx);
            return r2;
        }
    }
}
