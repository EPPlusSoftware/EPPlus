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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public class MultipleCharSeparatorHandler : SeparatorHandler
    {
        ITokenSeparatorProvider _tokenSeparatorProvider;
        INameValueProvider _nameValueProvider;

        public MultipleCharSeparatorHandler(INameValueProvider nameValueProvider)
            : this(new TokenSeparatorProvider(), nameValueProvider)
        {

        }
        public MultipleCharSeparatorHandler(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
            _nameValueProvider = nameValueProvider;
        }
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            // two operators in sequence could be "<=" or ">="
            if (IsPartOfMultipleCharSeparator(context, c))
            {
                var sOp = context.LastToken.Value.Value + c.ToString();
                var op = _tokenSeparatorProvider.Tokens[sOp];
                context.ReplaceLastToken(op);
                context.NewToken();
                return true;
            }
            if (c == ':')
            {
                HandleAddressSeparatorToken(c, tokenSeparator, context);
                return true;
            }
            return false;
        }

        private bool HandleAddressSeparatorToken(char c, Token tokenSeparator, TokenizerContext context)
        {
            if (context.LastToken != null && context.LastToken.Value.Value == ")")
            {
                context.AddToken(tokenSeparator);
            }
            else
            {
                if(_nameValueProvider.IsNamedValue(context.CurrentToken, context.Worksheet))
                {
                    var nameValue = _nameValueProvider.GetNamedValue(context.CurrentToken, context.Worksheet);
                    if(nameValue != null)
                    {
                        if(nameValue is IRangeInfo rangeInfo)
                        {
                            context.IsInDefinedNameAddress = true;
                            context.OverwriteCurrentToken(rangeInfo.Address.Address + ":");
                            return true;
                        }
                    }
                }
                context.AppendToCurrentToken(c);
            }
            return false;
        }

        private bool IsPartOfMultipleCharSeparator(TokenizerContext context, char c)
        {
            var lastTokenVal = string.Empty;
            if (!context.LastToken.HasValue) return false;
            var lastToken = context.LastToken.Value;
            lastTokenVal = lastToken.Value ?? string.Empty;
            return _tokenSeparatorProvider.IsOperator(lastTokenVal)
                && _tokenSeparatorProvider.IsPossibleLastPartOfMultipleCharOperator(c.ToString())
                && !context.CurrentTokenHasValue;
        }
    }
}
