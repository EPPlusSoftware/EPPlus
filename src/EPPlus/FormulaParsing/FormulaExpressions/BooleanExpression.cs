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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Data;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class BooleanExpression : Expression
    {
        private BooleanExpression _negatedExpression;
        internal BooleanExpression(string tokenValue, ParsingContext ctx) : base(ctx)
        {
            var value = bool.Parse(tokenValue);
            _cachedCompileResult = new CompileResult(value, DataType.Boolean);
            _negatedExpression = new BooleanExpression(ctx, this, !value);
        }
        internal BooleanExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
            _negatedExpression = new BooleanExpression(ctx, this, !((bool)result.ResultValue));
        }

        public BooleanExpression(ParsingContext ctx) : base(ctx)
        {
        }
        public BooleanExpression(ParsingContext ctx, BooleanExpression exp, bool negatedValue) : base(ctx)
        {
            _cachedCompileResult = new CompileResult(negatedValue,DataType.Boolean);
            _negatedExpression = exp;
        }

        internal override ExpressionType ExpressionType => ExpressionType.Boolean;

        public override CompileResult Compile()
        {
            return _cachedCompileResult;
        }
        public override Expression Negate()
        {
            return _negatedExpression;
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
    }
}
