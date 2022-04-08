/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class GroupExpression : Expression
    {
        public GroupExpression(bool isNegated)
            : this(isNegated, new ExpressionCompiler())
        {

        }

        public GroupExpression(bool isNegated, IExpressionCompiler expressionCompiler)
        {
            _expressionCompiler = expressionCompiler;
            _isNegated = isNegated;
        }

        private readonly IExpressionCompiler _expressionCompiler;
        private readonly bool _isNegated;


        public override CompileResult Compile()
        {
            var result =  _expressionCompiler.Compile(Children);
            if (result.IsNumeric && _isNegated)
            {
                return new CompileResult(result.ResultNumeric * -1, result.DataType);
            }
            return result;
        }

        public override bool IsGroupedExpression
        {
            get { return true; }
        }
    }
}
