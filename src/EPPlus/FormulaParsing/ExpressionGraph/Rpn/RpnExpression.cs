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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.RpnExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal enum RpnExpressionStatus : short
    {
        NoSet=0,
        CanCompile=1,
        IsAddress=2,
        OnExpressionList=4
    }
    internal abstract class RpnExpression
    {
        protected CompileResult _cachedCompileResult;
        internal Operators Operator;

        protected ParsingContext Context { get; private set; }
        internal abstract ExpressionType ExpressionType { get; }

        public RpnExpression(ParsingContext ctx)
        {
            Context = ctx;
        }
        public abstract CompileResult Compile();
        public abstract void Negate();
        internal abstract RpnExpressionStatus Status { get; set; }
    }
}
