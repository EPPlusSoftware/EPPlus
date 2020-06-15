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
    public class BracketHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if (context.IsInSheetName || context.IsInSheetName) return false;
            if (tokenSeparator.TokenTypeIsSet(TokenType.OpeningBracket))
            {
                context.AppendToCurrentToken(c);
                context.BracketCount++;
                return true;
            }
            if (tokenSeparator.TokenTypeIsSet(TokenType.ClosingBracket))
            {
                context.AppendToCurrentToken(c);
                context.BracketCount--;
                return true;
            }
            if (context.BracketCount > 0)
            {
                context.AppendToCurrentToken(c);
                return true;
            }
            return false;
        }
    }
}
