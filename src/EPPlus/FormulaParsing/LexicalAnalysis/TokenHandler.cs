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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenHandler : ITokenIndexProvider
    {
        public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider, INameValueProvider nameValueProvider)
            : this(context, tokenFactory, tokenProvider, new TokenSeparatorHandler(tokenProvider, nameValueProvider))
        {

        }
        public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider, TokenSeparatorHandler tokenSeparatorHandler)
        {
            _context = context;
            _tokenFactory = tokenFactory;
            _tokenProvider = tokenProvider;
            _tokenSeparatorHandler = tokenSeparatorHandler;
        }

        private readonly TokenizerContext _context;
        private readonly ITokenSeparatorProvider _tokenProvider;
        private readonly ITokenFactory _tokenFactory;
        private readonly TokenSeparatorHandler _tokenSeparatorHandler;
        private int _tokenIndex = -1;

        public string Worksheet { get; set; }

        public bool HasMore()
        {
            return _tokenIndex < (_context.FormulaChars.Length - 1);
        }

        public void Next()
        {
            _tokenIndex++;
            Handle();
        }

        private void Handle()
        {
            var c = _context.FormulaChars[_tokenIndex];
            Token tokenSeparator;
            if (CharIsTokenSeparator(c, out tokenSeparator))
            {
                if (_tokenSeparatorHandler.Handle(c, tokenSeparator, _context, this))
                {
                    return;
                }
                                              
                if (_context.CurrentTokenHasValue)
                {
                    if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
                    {
                        _context.AddToken(_tokenFactory.Create(_context.CurrentToken, TokenType.StringContent));
                    }
                    else
                    {
                        _context.AddToken(CreateToken(_context, Worksheet));
                    }


                    //If the a next token is an opening parantheses and the previous token is interpeted as an address or name, then the currenct token is a function
                    if (tokenSeparator.TokenTypeIsSet(TokenType.OpeningParenthesis) && (_context.LastToken.Value.TokenTypeIsSet(TokenType.ExcelAddress) || _context.LastToken.Value.TokenTypeIsSet(TokenType.NameValue)))
                    {
                        var newToken = _context.LastToken.Value.CloneWithNewTokenType(TokenType.Function);
                        _context.ReplaceLastToken(newToken);
                    }
                }
                if (TokenIsNegator(tokenSeparator.Value, _context))
                {
                    _context.AddToken(new Token("-", TokenType.Negator));
                    return;
                }
                _context.AddToken(tokenSeparator);
                _context.NewToken();
                return;
            }
            _context.AppendToCurrentToken(c);
        }

        private bool CharIsTokenSeparator(char c, out Token token)
        {
            var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
            token = result ? token = _tokenProvider.Tokens[c.ToString()] : default(Token);
            return result;
        }

        private static bool TokenIsNegator(string token, TokenizerContext context)
        {
            if (token != "-") return false;
            if (!context.LastToken.HasValue) return true;
            var t = context.LastToken.Value;
            
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

        private Token CreateToken(TokenizerContext context, string worksheet)
        {
            if (context.CurrentToken == "-")
            {
                if (context.LastToken == default(Token) && context.LastToken.Value.TokenTypeIsSet(TokenType.Operator))
                {
                    return new Token("-", TokenType.Negator);
                }
            }
            return _tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
        }

        int ITokenIndexProvider.Index
        {
            get { return _tokenIndex; }
        }


        void ITokenIndexProvider.MoveIndexPointerForward()
        {
            _tokenIndex++;
        }
    }
}
