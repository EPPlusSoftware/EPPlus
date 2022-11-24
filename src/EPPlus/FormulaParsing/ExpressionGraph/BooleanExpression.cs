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
    public class BooleanExpression : AtomicExpression
    {
        private bool? _precompiledValue;

        public BooleanExpression(string expression, ParsingContext ctx)
            : base(expression, ctx)
        {

        }

        public BooleanExpression(bool value, ParsingContext ctx)
            : base(value ? "true" : "false", ctx)
        {
            _precompiledValue = value;
        }
        internal override ExpressionType ExpressionType => ExpressionType.Boolean;
        public override CompileResult Compile()
        {
            var result = _precompiledValue ?? bool.Parse(ExpressionString);
            return new CompileResult(result, DataType.Boolean);
        }
    }
}
