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
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public class CompileStrategyFactory : ICompileStrategyFactory
    {
        public CompileStrategyFactory(ParsingContext ctx)
        {
            _context = ctx;
        }

        private readonly ParsingContext _context;

        public CompileStrategy Create(Expression expression)
        {
            if (expression.Operator.Operator == Operators.Concat)
            {
                return new StringConcatStrategy(expression, _context);
            }
            else
            {
                return new DefaultCompileStrategy(expression, _context);
            }
        }
    }
}
