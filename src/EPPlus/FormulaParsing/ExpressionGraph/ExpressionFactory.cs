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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionFactory : IExpressionFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
        {
            _excelDataProvider = excelDataProvider;
            _parsingContext = context;
        }


        public Expression Create(Token token)
        {
            if(token.TokenTypeIsSet(TokenType.Integer))
            {
                return new IntegerExpression(token.Value, token.IsNegated);
            }
            if (token.TokenTypeIsSet(TokenType.String))
            {
                return new StringExpression(token.Value);
            }
            if (token.TokenTypeIsSet(TokenType.Decimal))
            {
                return new DecimalExpression(token.Value, token.IsNegated);
            }
            if (token.TokenTypeIsSet(TokenType.Boolean))
            {
                return new BooleanExpression(token.Value);
            }
            if (token.TokenTypeIsSet(TokenType.ExcelAddress))
            {
                var exp = new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext, token.IsNegated);
                exp.HasCircularReference = token.TokenTypeIsSet(TokenType.CircularReference);
                return exp;
            }
            if (token.TokenTypeIsSet(TokenType.InvalidReference))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Ref));
            }
            if (token.TokenTypeIsSet(TokenType.NumericError))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Num));
            }
            if (token.TokenTypeIsSet(TokenType.ValueDataTypeError))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Value));
            }
            if (token.TokenTypeIsSet(TokenType.Null))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Null));
            }
            if (token.TokenTypeIsSet(TokenType.NameValue))
            {
                return new NamedValueExpression(token.Value, _parsingContext);
            }
            return new StringExpression(token.Value);
        }
    }
}
