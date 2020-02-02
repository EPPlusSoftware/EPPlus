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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing
{
    /// <summary>
    /// Postprocessor for a <see cref="TokenizerContext"/>
    /// </summary>
    public class TokenizerPostProcessor
    {
        public TokenizerPostProcessor(TokenizerContext context)
            : this(context, new TokenNavigator(context.Result))
        {

        }
        public TokenizerPostProcessor(TokenizerContext context, TokenNavigator navigator)
        {
            _context = context;
            _navigator = navigator;
        }

        private readonly TokenizerContext _context;
        private readonly TokenNavigator _navigator;
        private readonly Token PlusToken = TokenSeparatorProvider.Instance.GetToken("+").Value;
        private readonly Token MinusToken = TokenSeparatorProvider.Instance.GetToken("-").Value;

        /// <summary>
        /// Processes the <see cref="TokenizerContext"/>
        /// </summary>
        public void Process()
        {
            while (_navigator.HasNext())
            {
                var token = _navigator.CurrentToken;
                if(token.TokenTypeIsSet(TokenType.Unrecognized))
                {
                    HandleUnrecognizedToken();
                    break;
                }
                else if(token.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    HandleWorksheetNameToken();
                    break;
                }
                else if(token.TokenTypeIsSet(TokenType.Operator) || token.TokenTypeIsSet(TokenType.Negator))
                {
                    if (token.Value == "+" || token.Value == "-")
                        HandleNegators();
                }
                _navigator.MoveNext();
            }
        }

        private void ChangeTokenTypeOnCurrentToken(TokenType tokenType)
        {
            _context.ChangeTokenType(tokenType, _navigator.Index);
        }

        private void ChangeValueOnCurrentToken(string value)
        {
            _context.ChangeValue(value, _navigator.Index);
        }

        private void HandleNegators()
        {
            var token = _navigator.CurrentToken;
            //Remove '+' from start of formula and formula arguments
            if (token.Value == "+" && (!_navigator.HasPrev() || _navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.OpeningParenthesis) || _navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.Comma)))
            {
                RemoveTokenAndSetNegatorOperator();
                return;
            }

            var nextToken = _navigator.NextToken;
            if (nextToken.TokenTypeIsSet(TokenType.Operator) || nextToken.TokenTypeIsSet(TokenType.Negator))
            {
                // Remove leading '+' from operator combinations
                if (token.Value == "+" && (nextToken.Value == "+" || nextToken.Value == "-"))
                {
                    RemoveTokenAndSetNegatorOperator();
                }
                // Remove trailing '+' from a negator operation
                else if (token.Value == "-" && nextToken.Value == "+")
                {
                    RemoveTokenAndSetNegatorOperator(1);
                }
                // Convert double negator operation to positive declaration
                else if (token.Value == "-" && nextToken.Value == "-")
                {
                    _context.RemoveAt(_navigator.Index);
                    _context.Replace(_navigator.Index, PlusToken);
                    if (_navigator.Index > 0) _navigator.MoveIndex(-1);
                    HandleNegators();
                }
            }
        }

        private void HandleUnrecognizedToken()
        {
            if (_navigator.HasNext())
            {
                if (_navigator.NextToken.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    ChangeTokenTypeOnCurrentToken(TokenType.Function);
                }
                else
                {
                    ChangeTokenTypeOnCurrentToken(TokenType.NameValue);
                }
            }
            else
            {
                ChangeTokenTypeOnCurrentToken(TokenType.NameValue);
            }
        }


        private void HandleWorksheetNameToken()
        {
            // use this and the following three tokens
            var tokenType = _navigator.GetTokenAtRelativePosition(3).GetTokenTypeFlags();
            ChangeTokenTypeOnCurrentToken(tokenType);
            var sb = new StringBuilder();
            var nToRemove = 3;
            if (_navigator.NbrOfRemainingTokens < nToRemove)
            {
                ChangeTokenTypeOnCurrentToken(TokenType.InvalidReference);
                nToRemove = _navigator.NbrOfRemainingTokens;
            }
            else if (!_navigator.GetTokenAtRelativePosition(3).TokenTypeIsSet(TokenType.ExcelAddress) &&
                    !_navigator.GetTokenAtRelativePosition(3).TokenTypeIsSet(TokenType.ExcelAddressR1C1))
            {
                ChangeTokenTypeOnCurrentToken(TokenType.InvalidReference);
                nToRemove--;
            }
            else
            {
                for (var ix = 0; ix < 4; ix++)
                {
                    sb.Append(_navigator.GetTokenAtRelativePosition(ix).Value);
                }
            }
            ChangeValueOnCurrentToken(sb.ToString());
            for (var ix = 0; ix < nToRemove; ix++)
            {
                _context.RemoveAt(_navigator.Index + 1);
            }
        }

        private void SetNegatorOperator(int i)
        {
            var token = _context.Result[i];
            if (token.Value == "-" && i > 0 && (token.TokenTypeIsSet(TokenType.Operator) || token.TokenTypeIsSet(TokenType.Negator)))
            {
                if (TokenIsNegator(_context.Result[i - 1]))
                {
                    _context.Replace(i, new Token("-", TokenType.Negator));
                }
                else
                {
                    _context.Replace(i, MinusToken);
                }
            }
        }

        private bool TokenIsNegator(TokenizerContext context)
        {
            return TokenIsNegator(context.LastToken.Value);
        }
        private bool TokenIsNegator(Token t)
        {
            return t.TokenTypeIsSet(TokenType.Operator)
                        ||
                        t.TokenTypeIsSet(TokenType.OpeningParenthesis)
                        ||
                        t.TokenTypeIsSet(TokenType.Comma)
                        ||
                        t.TokenTypeIsSet(TokenType.SemiColon)
                        ||
                        t.TokenTypeIsSet(TokenType.OpeningEnumerable);
        }

        private void RemoveTokenAndSetNegatorOperator(int offset = 0)
        {
            _context.Result.RemoveAt(_navigator.Index + offset);
            SetNegatorOperator(_navigator.Index);
            if (_navigator.Index > 0) _navigator.MoveIndex(-1); ;
        }
    }
}
