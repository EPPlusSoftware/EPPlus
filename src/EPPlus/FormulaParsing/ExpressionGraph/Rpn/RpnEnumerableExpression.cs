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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    /// <summary>
    /// This expression represents a literal array where rows and cols are separated with comma and semicolon.
    /// </summary>
    internal class RpnEnumerableExpression : RpnExpression
    {
        
        private readonly IRangeInfo _range;
        private bool _isNegated;

        internal RpnEnumerableExpression(CompileResult result, ParsingContext ctx)
            : base(ctx)
        {
            _cachedCompileResult = result;
        }
        internal RpnEnumerableExpression(IRangeInfo range, ParsingContext ctx)
            : base(ctx)
        {
            _range = range;
        }
        internal override ExpressionType ExpressionType => ExpressionType.Enumerable;
        /// <summary>
        /// Compiles the expression into a <see cref="CompileResult"/>
        /// </summary>
        /// <returns></returns>
        public override CompileResult Compile()
        {
            if(_cachedCompileResult==null)
            {
                _cachedCompileResult = new CompileResult(_range, DataType.ExcelRange);
            }
            return _cachedCompileResult;
        }

        public override void Negate()
        {
            _isNegated = !_isNegated;
        }
        internal override RpnExpressionStatus Status
        {
            get;
            set;
        } = RpnExpressionStatus.CanCompile;
    }
}
