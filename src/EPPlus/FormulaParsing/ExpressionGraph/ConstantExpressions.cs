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
    public static class ConstantExpressions
    {
        public static Expression Percent
        {
            get { return new ConstantExpression("Percent", () => new CompileResult(0.01, DataType.Decimal), ParsingContext.Create()); }
        }
    }

    public class ConstantExpression : AtomicExpression
    {
        private readonly Func<CompileResult> _factoryMethod;

        public ConstantExpression(string title, Func<CompileResult> factoryMethod, ParsingContext ctx)
            : base(title, ctx)
        {
            _factoryMethod = factoryMethod;
        }

        internal override ExpressionType ExpressionType => ExpressionType.Constant;

        public override CompileResult Compile()
        {
            return _factoryMethod();
        }
    }
}
