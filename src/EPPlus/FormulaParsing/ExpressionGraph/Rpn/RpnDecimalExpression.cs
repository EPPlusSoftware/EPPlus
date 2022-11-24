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
    internal class RpnDecimalExpression : RpnExpression
    {
        internal RpnDecimalExpression(string tokenValue, ParsingContext ctx) : base(ctx)
        {
            var value = double.Parse(tokenValue, CultureInfo.InvariantCulture);
            _cachedCompileResult = new CompileResult(value, DataType.Decimal);
        }
        internal RpnDecimalExpression(CompileResult result, ParsingContext ctx) : base(ctx)
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
