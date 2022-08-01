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
    /// <summary>
    /// Handles parsing of worksheet names
    /// </summary>
    internal class SheetnameHandler : SeparatorHandler
    {
        /// <summary>
        /// Handles characters and appends them to the sheetname
        /// </summary>
        /// <param name="c"></param>
        /// <param name="tokenSeparator"></param>
        /// <param name="context"></param>
        /// <param name="tokenIndexProvider"></param>
        /// <returns></returns>
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if (context.IsInSheetName)
            {
                if (IsDoubleQuote(tokenSeparator, tokenIndexProvider.Index, context))
                {
                    tokenIndexProvider.MoveIndexPointerForward();
                    context.AppendToCurrentToken(c);
                    return true;
                }
                else if(IsDoubleSingleQuote(tokenSeparator, tokenIndexProvider.Index, context))
                {
                    tokenIndexProvider.MoveIndexPointerForward();
                    // double single quotes inside a sheet name should be preserved
                    context.AppendToCurrentToken(c);
                    context.AppendToCurrentToken(c);
                    return true;
                }
                if (!tokenSeparator.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    context.AppendToCurrentToken(c);
                    return true;
                }
            }

            if (tokenSeparator.TokenTypeIsSet(TokenType.WorksheetName))
            {                
                if (context.LastToken != null && context.LastToken.Value.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    if (context.CurrentToken.StartsWith("!") && context.CurrentToken.EndsWith(":"))
                    {                        
                        context.AddToken(new Token(context.CurrentToken.Substring(0,context.CurrentToken.Length-1), TokenType.ExcelAddress));
                        context.AddToken(new Token(":", TokenType.Colon));
                        context.NewToken();
                    }
                    else
                    {
                        context.AddToken(!context.CurrentTokenHasValue
                            ? new Token(string.Empty, TokenType.WorksheetNameContent)
                            : new Token(context.CurrentToken, TokenType.WorksheetNameContent));
                    }
                }
                else if(context.CurrentToken.EndsWith(":"))
                {
                    context.AddToken(new Token(context.CurrentToken.Substring(0, context.CurrentToken.Length - 1), TokenType.ExcelAddress));
                    context.AddToken(new Token(":", TokenType.Colon));
                    context.NewToken();
                }
                if (context.CurrentToken.StartsWith("[") &&
                   context.CurrentToken.EndsWith("]"))
                {
                    context.AddToken(new Token(context.CurrentToken + "'", TokenType.WorksheetName)); //Append current token, as this can be an external reference index e.g [1]
                }
                else
                {
                    context.AddToken(new Token("'", TokenType.WorksheetName)); //Append current token, as this can be an external reference index e.g [1]
                }
                context.ToggleIsInSheetName();
                context.NewToken();
                return true;
            }
            return false;
        }
    }
}
