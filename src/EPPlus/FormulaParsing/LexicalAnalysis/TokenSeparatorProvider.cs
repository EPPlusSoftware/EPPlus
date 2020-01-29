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
using System.Threading;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenSeparatorProvider : ITokenSeparatorProvider
    {
       private static readonly Dictionary<string, Token> _tokens;

        static TokenSeparatorProvider()
        {
            _tokens = new Dictionary<string, Token>();
            _tokens.Add("+", new Token("+", TokenType.Operator));
            _tokens.Add("-", new Token("-", TokenType.Operator));
            _tokens.Add("*", new Token("*", TokenType.Operator));
            _tokens.Add("/", new Token("/", TokenType.Operator));
            _tokens.Add("^", new Token("^", TokenType.Operator));
            _tokens.Add("&", new Token("&", TokenType.Operator));
            _tokens.Add(">", new Token(">", TokenType.Operator));
            _tokens.Add("<", new Token("<", TokenType.Operator));
            _tokens.Add("=", new Token("=", TokenType.Operator));
            _tokens.Add("<=", new Token("<=", TokenType.Operator));
            _tokens.Add(">=", new Token(">=", TokenType.Operator));
            _tokens.Add("<>", new Token("<>", TokenType.Operator));
            _tokens.Add("(", new Token("(", TokenType.OpeningParenthesis));
            _tokens.Add(")", new Token(")", TokenType.ClosingParenthesis));
            _tokens.Add("{", new Token("{", TokenType.OpeningEnumerable));
            _tokens.Add("}", new Token("}", TokenType.ClosingEnumerable));
            _tokens.Add("'", new Token("'", TokenType.WorksheetName));
            _tokens.Add("\"", new Token("\"", TokenType.String));
            _tokens.Add(",", new Token(",", TokenType.Comma));
            _tokens.Add(";", new Token(";", TokenType.SemiColon));
            _tokens.Add("[", new Token("[", TokenType.OpeningBracket));
            _tokens.Add("]", new Token("]", TokenType.ClosingBracket));
            _tokens.Add("%", new Token("%", TokenType.Percent));
        }

        IDictionary<string, Token> ITokenSeparatorProvider.Tokens
        {
            get { return _tokens; }
        }

        /// <summary>
        /// Returns true if the item is an operator, otherwise false.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool IsOperator(string item)
        {
            Token token;
            if (_tokens.TryGetValue(item, out token))
            {
                if (token.TokenType == TokenType.Operator)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Returns true if the <paramref name="part"/> could be part of a multichar operator, such as != or &lt;&gt;
        /// </summary>
        /// <param name="part"></param>
        /// <returns></returns>
        public bool IsPossibleLastPartOfMultipleCharOperator(string part)
        {
            return part == "=" || part == ">";
        }

        /// <summary>
        /// Returns a separator <see cref="Token"/> by its string representation.
        /// </summary>
        /// <param name="candidate">The separator candidate</param>
        /// <returns>A <see cref="Token"/> instance or null/default(Token?)</returns>
        public Token? GetToken(string candidate)
        {
            if (_tokens.ContainsKey(candidate)) return _tokens[candidate];
            return default(Token?);
        }

        /// <summary>
        /// Instance of the <see cref="ITokenSeparatorProvider"/>
        /// </summary>
        public static ITokenSeparatorProvider Instance { get; } = new TokenSeparatorProvider();
    }
}
