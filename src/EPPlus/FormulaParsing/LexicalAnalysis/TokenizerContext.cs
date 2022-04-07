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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Responsible for handling tokens during the tokenizing process.
    /// </summary>
    internal class TokenizerContext
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="formula">The formula to tokenize</param>
        /// <param name="worksheet">Worksheet name, if applicable</param>
        /// <param name="tokenFactory">A <see cref="ITokenFactory"/> instance</param>
        public TokenizerContext(string formula, string worksheet, ITokenFactory tokenFactory)
        {
            if (!string.IsNullOrEmpty(formula))
            {
                FormulaChars = formula.ToArray();
            }
            Require.That(tokenFactory).IsNotNull();
            _result = new List<Token>();
            _currentToken = new StringBuilder();
            _worksheet = worksheet;
            _tokenFactory = tokenFactory;
        }

        private readonly List<Token> _result;
        private StringBuilder _currentToken;
        private readonly ITokenFactory _tokenFactory;
        private readonly string _worksheet;

        /// <summary>
        /// The formula split into a character array
        /// </summary>
        public char[] FormulaChars
        {
            get; private set;
        }

        public TokenHandler CreateHandler(INameValueProvider nameValueProvider)
        {
            var handler = new TokenHandler(this, _tokenFactory, TokenSeparatorProvider.Instance, nameValueProvider);
            handler.Worksheet = _worksheet;
            return handler;
        }

        /// <summary>
        /// The tokens created
        /// </summary>
        public IList<Token> Result
        {
            get { return _result; }
        }

        internal string Worksheet
        {
            get { return _worksheet;}
        }

        /// <summary>
        /// Returns the token before the requested index
        /// </summary>
        /// <param name="index">The requested index</param>
        /// <returns>The <see cref="Token"/> at the requested position</returns>
        public Token GetTokenBeforeIndex(int index)
        {
            if (index < 1 || index > _result.Count - 1) throw new IndexOutOfRangeException("Index was out of range of the token array");
            return _result[index - 1];
        }

        /// <summary>
        /// Returns the token after the requested index
        /// </summary>
        /// <param name="index">The requested index</param>
        /// <returns>The <see cref="Token"/> at the requested position</returns>
        public Token GetNextTokenAfterIndex(int index)
        {
            if (index < 0 || index > _result.Count - 2) throw new IndexOutOfRangeException("Index was out of range of the token array");
            return _result[index + 1];
        }

        private Token CreateToken(string worksheet)
        {
            if (CurrentToken == "-")
            {
                if (LastToken == null && LastToken.Value.TokenTypeIsSet(TokenType.Operator))
                {
                    return new Token("-", TokenType.Negator);
                }
            }
            return _tokenFactory.Create(Result, CurrentToken, worksheet);
        }

        internal void OverwriteCurrentToken(string token)
        {
            _currentToken = new StringBuilder(token);
        }

        public void PostProcess()
        {
            if (CurrentTokenHasValue)
            {
                AddToken(CreateToken(_worksheet));
            }

            var postProcessor = new TokenizerPostProcessor(this);
            postProcessor.Process();
        }

        /// <summary>
        /// Replaces a token at the requested <paramref name="index"/>
        /// </summary>
        /// <param name="index">0-based index of the requested position</param>
        /// <param name="newValue">The new <see cref="Token"/></param>
        public void Replace(int index, Token newValue)
        {
            _result[index] = newValue;
        }

        /// <summary>
        /// Removes the token at the requested <see cref="Token"/>
        /// </summary>
        /// <param name="index">0-based index of the requested position</param>
        public void RemoveAt(int index)
        {
            _result.RemoveAt(index);
        }

        /// <summary>
        /// Returns true if the current position is inside a string, otherwise false.
        /// </summary>
        public bool IsInString
        {
            get;
            private set;
        }

        /// <summary>
        /// Returns true if the current position is inside a sheetname, otherwise false.
        /// </summary>
        public bool IsInSheetName
        {
            get;
            private set;
        }

        /// <summary>
        /// Toggles the IsInString state.
        /// </summary>
        public void ToggleIsInString()
        {
            IsInString = !IsInString;
        }

        /// <summary>
        /// Toggles the IsInSheetName state
        /// </summary>
        public void ToggleIsInSheetName()
        {
            IsInSheetName = !IsInSheetName;
        }

        internal int BracketCount
        {
            get;
            set;
        }

        internal bool IsInDefinedNameAddress
        {
            get;
            set;
        }

        /// <summary>
        /// Returns the current
        /// </summary>
        public string CurrentToken
        {
            get { return _currentToken.ToString(); }
        }

        public bool CurrentTokenHasValue
        {
            get { return !string.IsNullOrEmpty(IsInString ? CurrentToken : CurrentToken.Trim()); }
        }

        public void NewToken()
        {
            _currentToken = new StringBuilder();
        }

        public void AddToken(Token token)
        {
            _result.Add(token);
        }

        public void AppendToCurrentToken(char c)
        {
            _currentToken.Append(c.ToString());
        }

        public void AppendToLastToken(string stringToAppend)
        {
            var token = _result.Last();
            var newVal = token.Value += stringToAppend;
            var newToken = token.CloneWithNewValue(newVal);
            ReplaceLastToken(newToken);
        }

        /// <summary>
        /// Changes <see cref="TokenType"/> of the current token.
        /// </summary>
        /// <param name="tokenType">The new <see cref="TokenType"/></param>
        /// <param name="index">Index of the token to change</param>
        public void ChangeTokenType(TokenType tokenType, int index)
        {
            _result[index] = _result[index].CloneWithNewTokenType(tokenType);
        }

        /// <summary>
        /// Changes the value of the current token
        /// </summary>
        /// <param name="val"></param>
        /// <param name="index">Index of the token to change</param>
        public void ChangeValue(string val, int index)
        {
            _result[index] = _result[index].CloneWithNewValue(val);
        }

        /// <summary>
        /// Changes the <see cref="TokenType"/> of the last token in the result.
        /// </summary>
        /// <param name="type"></param>
        public void SetLastTokenType(TokenType type)
        {
            var newToken = _result.Last().CloneWithNewTokenType(type);
            ReplaceLastToken(newToken);
        }

        /// <summary>
        /// Replaces the last token of the result with the <paramref name="newToken"/>
        /// </summary>
        /// <param name="newToken">The new token</param>
        public void ReplaceLastToken(Token newToken)
        {
            var count = _result.Count;
            if (count > 0)
            {
                _result.RemoveAt(count - 1);   
            }
            _result.Add(newToken);
        }

        /// <summary>
        /// Returns the last token of the result, if empty null/default(Token?) will be returned.
        /// </summary>
        public Token? LastToken
        {
            get { return _result.Count > 0 ? _result.Last() : default(Token?); }
        }

    }
}
