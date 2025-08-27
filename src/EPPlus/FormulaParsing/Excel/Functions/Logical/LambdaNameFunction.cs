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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.Utils.Formula;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    internal class LambdaNameFunction : ExcelFunction
    {
        public LambdaNameFunction(string formula, VariableStorageScope scope)
        {
            _formula = formula;
            _scope = scope;
        }

        private readonly string _formula;
        private readonly VariableStorageScope _scope;
        private RpnFormula _rpnFormula;

        public override int ArgumentMinLength => 0;

        public override bool IsVolatile => true;

        internal override void SetRpnFormula(RpnFormula formula)
        {
            _rpnFormula = formula;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext ctx)
        {
            var elt = LambdaExtractor.GetLambdaTokens(_formula);
            var rpn2 = new RpnFormula(ctx.CurrentWorksheet, ctx.CurrentCell.Row, ctx.CurrentCell.Column, ctx.VariableStorage);
            rpn2.ExpressionStack = _rpnFormula.ExpressionStack;
            rpn2.FunctionStack = _rpnFormula.FunctionStack;
            rpn2.IgnoreCaching = true;
            var clc = new LambdaCalculator(elt.LambdaTokens.Tokens, _scope, rpn2);
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
            return clc.Execute(ctx);
        }
    }
}
