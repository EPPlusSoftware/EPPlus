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
    internal class TokenizerPostProcessor
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
            var hasColon = false;
            while (_navigator.HasNext())
            {
                var token = _navigator.CurrentToken;
                if (token.TokenTypeIsSet(TokenType.Unrecognized))
                {
                    HandleUnrecognizedToken();
                }
                else if (token.TokenTypeIsSet(TokenType.Colon))
                {
                    HandleColon();
                    hasColon = true;
                }
                else if (token.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    HandleWorksheetNameToken();
                }
                else if (token.TokenTypeIsSet(TokenType.Operator) || token.TokenTypeIsSet(TokenType.Negator))
                {
                    if (token.Value == "+" || token.Value == "-")
                        HandleNegators();
                }
                _navigator.MoveNext();
            }

            if (hasColon)
            {
                _navigator.MoveIndex(-_navigator.Index);
                while (_navigator.HasNext())
                {
                    var token = _navigator.CurrentToken;
                    if (token.TokenTypeIsSet(TokenType.Colon) && _context.Result.Count > _navigator.Index + 1)
                    {
                        if (_navigator.PreviousToken != null && _navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.ExcelAddress) &&
                           _navigator.NextToken.TokenTypeIsSet(TokenType.ExcelAddress))
                        {
                            var newToken= _navigator.PreviousToken.Value.Value+":"+ _navigator.NextToken.Value;
                            _context.Result[_navigator.Index-1] = new Token(newToken, TokenType.ExcelAddress);
                            _context.RemoveAt(_navigator.Index);
                            _context.RemoveAt(_navigator.Index);
                            _navigator.MoveIndex(-1);
                        }
                    }
                    _navigator.MoveNext();
                }
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

        private bool IsOffsetFunctionToken(Token token)
        {
            return token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset";
        }

        private void HandleColon()
        {
            var prevToken = _navigator.GetTokenAtRelativePosition(-1);
            var nextToken = _navigator.GetTokenAtRelativePosition(1);
            if (prevToken.TokenTypeIsSet(TokenType.ClosingParenthesis))
            {
                // Previous expression should be an OFFSET function
                var index = 0;
                var openedParenthesis = 0;
                var closedParethesis = 0;
                while(openedParenthesis == 0 || openedParenthesis > closedParethesis)
                {
                    index--;
                    var token = _navigator.GetTokenAtRelativePosition(index);
                    if (token.TokenTypeIsSet(TokenType.ClosingParenthesis))
                        openedParenthesis++;
                    else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                        closedParethesis++;
                }
                var offsetCandidate = _navigator.GetTokenAtRelativePosition(--index);
                if(IsOffsetFunctionToken(offsetCandidate))
                {
                    _context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, _navigator.Index + index);
                    if (nextToken.TokenTypeIsSet(TokenType.ExcelAddress))
                    {
                        // OFFSET:A1
                        _context.ChangeTokenType(TokenType.ExcelAddress | TokenType.RangeOffset, _navigator.Index + 1);
                    }
                    else if(IsOffsetFunctionToken(nextToken))
                    {
                        // OFFSET:OFFSET
                        _context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, _navigator.Index + 1);
                    }
                }
            }
            else if(prevToken.TokenTypeIsSet(TokenType.ExcelAddress) && IsOffsetFunctionToken(nextToken))
            {
                // A1: OFFSET
                _context.ChangeTokenType(TokenType.ExcelAddress | TokenType.RangeOffset, _navigator.Index - 1);
                _context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, _navigator.Index + 1);
            }
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
                    _context.ChangeTokenType(TokenType.Negator, _navigator.Index);
                    _navigator.MoveIndex(1);
                    _context.ChangeTokenType(TokenType.Negator, _navigator.Index);
                    /*
                    _context.RemoveAt(_navigator.Index);
                    _context.Replace(_navigator.Index, PlusToken);
                    if (_navigator.Index > 0) _navigator.MoveIndex(-1);
                    HandleNegators();
                    */
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
            var relativeToken = _navigator.GetTokenAtRelativePosition(3);
            var tokenType = relativeToken.GetTokenTypeFlags();
            ChangeTokenTypeOnCurrentToken(tokenType);
            var sb = new StringBuilder();
            var nToRemove = 3;
            if (_navigator.NbrOfRemainingTokens < nToRemove)
            {
                ChangeTokenTypeOnCurrentToken(TokenType.InvalidReference);
                nToRemove = _navigator.NbrOfRemainingTokens;
            }
            if (relativeToken.TokenTypeIsSet(TokenType.Comma) ||
               relativeToken.TokenTypeIsSet(TokenType.ClosingParenthesis))
            {
                for (var ix = 0; ix < 3; ix++)
                {
                    sb.Append(_navigator.GetTokenAtRelativePosition(ix).Value);
                }
                ChangeTokenTypeOnCurrentToken(TokenType.ExcelAddress);
                nToRemove = 2;
            }
            else if (!relativeToken.TokenTypeIsSet(TokenType.ExcelAddress) &&
                     !relativeToken.TokenTypeIsSet(TokenType.ExcelAddressR1C1) &&
                     !relativeToken.TokenTypeIsSet(TokenType.NameValue) &&
                     !relativeToken.TokenTypeIsSet(TokenType.InvalidReference))
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
