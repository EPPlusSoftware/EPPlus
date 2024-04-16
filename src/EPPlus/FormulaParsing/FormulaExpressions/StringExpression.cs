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
    internal class StringExpression : Expression
    {
        internal StringExpression(string tokenValue, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = new CompileResult(tokenValue.Substring(1,tokenValue.Length-2).Replace("\"\"", "\""), DataType.String); //Remove double quotes and 
        }
        internal StringExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
        }

        internal override ExpressionType ExpressionType => ExpressionType.String;

        public override CompileResult Compile()
        {
            return _cachedCompileResult;
        }
        public override Expression Negate()
        {
            var cr = _cachedCompileResult.Negate();
            if(cr.DataType == DataType.Decimal)
            {
                return new DecimalExpression(cr, Context);
            }
            else
            {
                return ErrorExpression.ValueError;
            }
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
    }
}


