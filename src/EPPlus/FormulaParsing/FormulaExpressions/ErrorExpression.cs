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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Globalization;
namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class ErrorExpression : Expression
    {
        internal static ErrorExpression RefError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Ref), null);
        internal static ErrorExpression ValueError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Value), null);
        internal static ErrorExpression NaError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.NA), null);
        internal static ErrorExpression NameError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Name), null);
        internal static ErrorExpression NumError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Num), null);
        internal static ErrorExpression NullError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Null), null);
        internal static ErrorExpression Div0Error => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Div0), null);
        internal static ErrorExpression CalcError => new ErrorExpression(CompileResult.GetErrorResult(eErrorType.Calc), null);

        internal ErrorExpression(string tokenValue, ParsingContext ctx) : base(ctx)
        {
            var value = ExcelErrorValue.Parse(tokenValue);
            _cachedCompileResult = new CompileResult(value, DataType.ExcelError);
        }
        internal ErrorExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
        }

        internal override ExpressionType ExpressionType => ExpressionType.Decimal;

        public override CompileResult Compile()
        {
            return _cachedCompileResult;
        }
        public override Expression Negate()
        {
            return this;
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
    }
}
