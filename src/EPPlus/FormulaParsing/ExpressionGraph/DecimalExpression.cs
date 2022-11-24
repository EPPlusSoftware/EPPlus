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
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class DecimalExpression : AtomicExpression
    {
        private double? _compiledValue;
        private bool _negate;

        public DecimalExpression(string expression, ParsingContext ctx)
            : this(expression, false, ctx)
        {
            
        }

        public DecimalExpression(string expression, bool negate, ParsingContext ctx)
            : base(expression, ctx)
        {
            _negate = negate;
        }

        public DecimalExpression(double compiledValue, ParsingContext ctx)
            : base(compiledValue.ToString(CultureInfo.InvariantCulture), ctx)
        {
            _compiledValue = compiledValue;
        }
        internal override ExpressionType ExpressionType => ExpressionType.Decimal;
        public override CompileResult Compile()
        {
            double result = _compiledValue ?? double.Parse(ExpressionString, CultureInfo.InvariantCulture);
            result = _negate ? result * -1 : result;
            return new CompileResult(result, DataType.Decimal);
        }
    }
}
