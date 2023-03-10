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
namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal class RpnErrorExpression : RpnExpression
    {
        internal static RpnErrorExpression RefError => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.Ref), null);
        internal static RpnErrorExpression ValueError => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.Value), null);
        internal static RpnErrorExpression NaError => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.NA), null);
        internal static RpnErrorExpression NameError => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.Name), null);
        internal static RpnErrorExpression NumError => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.Num), null);
        internal static RpnErrorExpression Div0Error => new RpnErrorExpression(CompileResult.GetErrorResult(eErrorType.Div0), null);
        internal RpnErrorExpression(string tokenValue, ParsingContext ctx) : base(ctx)
        {
            var value = ExcelErrorValue.Parse(tokenValue);
            _cachedCompileResult = new CompileResult(value, DataType.ExcelError);
        }
        internal RpnErrorExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
        }

        internal override ExpressionType ExpressionType => ExpressionType.Decimal;

        public override CompileResult Compile()
        {
            return _cachedCompileResult;
        }
        public override void Negate()
        {
            _cachedCompileResult.Negate();
        }
        internal override RpnExpressionStatus Status
        {
            get;
            set;
        } = RpnExpressionStatus.CanCompile;
    }
}
