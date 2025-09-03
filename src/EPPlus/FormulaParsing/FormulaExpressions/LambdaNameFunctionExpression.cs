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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Represents a LAMBDA function stored as a Name in the workbook.
    /// </summary>
    internal class LambdaNameFunctionExpression : LambdaFunctionExpression
    {
        internal LambdaNameFunctionExpression(string functionName, string formula, ParsingContext ctx, int pos) : base(functionName, ctx, pos)
        {
            var scope = ctx.VariableStorage.GetById(VariableScopeId);
            _function = new LambdaNameFunction(formula, scope);
        }

        internal override void SetRpnFormula(RpnFormula formula)
        {
           _function.SetRpnFormula(formula);
        }

        public override CompileResult Compile()
        {
            return base.Compile();
        }

        internal override void OnDispose()
        {
            base.OnDispose();
        }

    }
}
