/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2022         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline;
using System.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal class RpnFunctionArgumentExpression : RpnExpression
    {
        bool _negate=false;
        internal int _startPos, _endPos;
        internal RpnFunctionArgumentExpression(ParsingContext ctx, int startPos) : base(ctx)
        {
            _startPos = startPos;
        }
        internal override ExpressionType ExpressionType => ExpressionType.Function;
        public override void Negate()
        {
            _negate = !_negate;
        }
        public override CompileResult Compile()
        {
            return new CompileResult(0, DataType.Empty);
        }
        private RpnExpressionStatus _status= RpnExpressionStatus.FunctionArgument;
        internal override RpnExpressionStatus Status
        {
            get
            {
                return _status;
            }
            set
            {
                _status = value;
            }
        }
    }

}
