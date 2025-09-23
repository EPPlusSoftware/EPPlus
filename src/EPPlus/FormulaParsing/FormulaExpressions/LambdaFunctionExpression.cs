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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// This expression handles the initial execution of a LAMBDA function.
    /// </summary>
    internal class LambdaFunctionExpression : VariableFunctionExpression
    {
        internal override bool IsLambda => true;
        internal override bool HandlesVariables => true;
        private List<Token> _lambdaTokens;
        
        internal LambdaFunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {

        }

        internal LambdaFunctionExpression(string tokenValue, ParsingContext ctx, int pos, bool addVariableScope) : base(tokenValue, ctx, pos, addVariableScope)
        {

        }

        internal override void SetRpnFormula(RpnFormula formula)
        {
            _function.SetRpnFormula(formula);
        }
    }
}
