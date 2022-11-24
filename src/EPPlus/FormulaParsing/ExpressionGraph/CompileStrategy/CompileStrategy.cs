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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public abstract class CompileStrategy
    {
        protected readonly Expression _expression;
        protected ParsingContext Context { get; private set; }

        public CompileStrategy(Expression expression, ParsingContext ctx)
        {
            _expression = expression;
            Context = ctx;
        }

        public abstract Expression Compile(IList<Expression> expressions, int index);
    }
}
