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
    internal class ExpressionFactory : IExpressionFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
        {
            _excelDataProvider = excelDataProvider;
            _parsingContext = context;
        }


        public Expression Create(Token token, ref FormulaAddressBase addressInfo, Expression parent)
        {
            if(token.TokenTypeIsSet(TokenType.Integer))
            {
                return new IntegerExpression(token.Value, token.IsNegated, _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.String))
            {
                return new StringExpression(token.Value, _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.Decimal))
            {
                return new DecimalExpression(token.Value, token.IsNegated, _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.Boolean))
            {
                return new BooleanExpression(token.Value, _parsingContext);
            }
            if(token.TokenTypeIsSet(TokenType.CellAddress))
            {
                return new CellAddressExpression(token, _parsingContext, ref addressInfo) { _parent = parent };
            }
            if((token.TokenTypeIsSet(TokenType.ClosingBracket) && addressInfo is FormulaTableAddress ti))
            {
                return new TableAddressExpression(_parsingContext, ti) { _parent = parent };
            }
            if (token.TokenTypeIsSet(TokenType.InvalidReference))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Ref), _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.NumericError))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Num), _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.ValueDataTypeError))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Value), _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.Null))
            {
                return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Null), _parsingContext);
            }
            if (token.TokenTypeIsSet(TokenType.NameValue))
            {
                return new NamedValueExpression(token.Value, _parsingContext, ref addressInfo) { _parent = parent };
            }
            return new StringExpression(token.Value, _parsingContext);
        }
    }
}
