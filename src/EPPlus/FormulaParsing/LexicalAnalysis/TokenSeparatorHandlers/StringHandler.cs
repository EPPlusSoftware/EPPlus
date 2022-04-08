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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    internal class StringHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if(context.IsInString)
            { 
                if (IsDoubleQuote(tokenSeparator, tokenIndexProvider.Index, context))
                {
                    tokenIndexProvider.MoveIndexPointerForward();
                    context.AppendToCurrentToken(c);
                    return true;
                }
                if (!tokenSeparator.TokenTypeIsSet(TokenType.String))
                {
                    context.AppendToCurrentToken(c);
                    return true;
                }
            }

            if (tokenSeparator.TokenTypeIsSet(TokenType.String))
            {
                if (context.LastToken != null && context.LastToken.Value.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    context.AppendToCurrentToken(c);
                    context.ToggleIsInString();
                    context.AddToken(tokenSeparator);
                    context.NewToken();
                    return true;
                }
                if (context.LastToken != null && context.LastToken.Value.TokenTypeIsSet(TokenType.String))
                {
                    context.AddToken(!context.CurrentTokenHasValue
                        ? new Token(string.Empty, TokenType.StringContent)
                        : new Token(context.CurrentToken, TokenType.StringContent));
                }
                context.AddToken(new Token("\"", TokenType.String));
                context.ToggleIsInString();
                context.NewToken();
                return true;
            }
            return false;
        }
    }
}
