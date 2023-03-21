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
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [Flags]
    internal enum ExpressionStatus : short
    {
        NoSet = 0,
        CanCompile = 1,
        IsAddress = 2,
        OnExpressionList = 4,
        FunctionArgument = 8
    }
    internal class EmptyExpression : Expression
    {
        internal override ExpressionType ExpressionType => ExpressionType.Empty;
        public override CompileResult Compile()
        {
            return CompileResult.Empty;
        }
        internal override ExpressionStatus Status { get; set; }
    }
    public abstract class Expression
    {
        protected CompileResult _cachedCompileResult;
        internal Operators Operator;
        internal static EmptyExpression Empty=new EmptyExpression();

        protected ParsingContext Context { get; private set; }
        internal abstract ExpressionType ExpressionType { get; }
        internal Expression()
        {
        }
        public Expression(ParsingContext ctx)
        {
            Context = ctx;
        }
        public abstract CompileResult Compile();
        public virtual void Negate()
        {

        }

        internal virtual Expression CloneWithOffset(int row, int col)
        {
            return this;
        }

        internal abstract ExpressionStatus Status { get; set; }
        public virtual FormulaRangeAddress GetAddress() { return null; }

        internal virtual void MergeAddress(string address)
        {

        }
    }
}
