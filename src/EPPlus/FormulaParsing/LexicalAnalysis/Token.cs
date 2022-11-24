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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Represents a character in a formula
    /// </summary>
    public struct Token
    {
        const ushort IS_NEGATED=0x0001;
        internal Token(TokenType tokenType)            
        {
            TokenType = tokenType;
            Value = null;
            _flags = 0;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="token">The formula character</param>
        /// <param name="tokenType">The <see cref="LexicalAnalysis.TokenType"/></param>
        public Token(string token, TokenType tokenType)
            : this(token, tokenType, false)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="token">The formula character</param>
        /// <param name="tokenType">The <see cref="LexicalAnalysis.TokenType"></see></param>
        /// <param name="isNegated"></param>
        public Token(string token, TokenType tokenType, bool isNegated)
        {
            Value = token;
            TokenType = tokenType;
            _flags = (ushort)(isNegated ? IS_NEGATED : 0);
        }

        /// <summary>
        /// The formula character
        /// </summary>
        public string Value;

        internal TokenType TokenType;
        internal ushort _flags;
        /// <summary>
        /// Indicates whether a numeric value should be negated when compiled
        /// </summary>
        public bool IsNegated 
        { 
            get
            {
                return (_flags & IS_NEGATED) > 0;
            }
        }

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <returns></returns>
        public static bool operator ==(Token t1, Token t2)
        {
            return t1.AreEqualTo(t2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <returns></returns>
        public static bool operator !=(Token t1, Token t2)
        {
            return !(t1.AreEqualTo(t2));
        }

        /// <summary>
        /// Overrides object.Equals with no behavioural change
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        /// <summary>
        /// Overrides object.GetHashCode with no behavioural change
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        /// Return if the supplied <paramref name="tokenType"/> is set on this token.
        /// </summary>
        /// <param name="tokenType">The <see cref="LexicalAnalysis.TokenType"></see> to check</param>
        /// <returns>True if the token is set, otherwirse false</returns>
        public bool TokenTypeIsSet(TokenType tokenType)
        {
            return (TokenType & tokenType) == tokenType;
        }

        public bool AreEqualTo(Token otherToken)
        {
            return (GetTokenTypeFlags() == otherToken.GetTokenTypeFlags() && Value == otherToken.Value);
        }

        internal TokenType GetTokenTypeFlags()
        {
            return TokenType;
        }

        /// <summary>
        /// Clones the token with a new <see cref="LexicalAnalysis.TokenType"/> set.
        /// </summary>
        /// <param name="tokenType">The new TokenType</param>
        /// <returns>A cloned Token</returns>
        internal Token CloneWithNewTokenType(TokenType tokenType)
        {
            return new Token(Value, tokenType, IsNegated);
        }

        /// <summary>
        /// Clones the token with a new value set.
        /// </summary>
        /// <param name="val">The new value</param>
        /// <returns>A cloned Token</returns>
        internal Token CloneWithNewValue(string val)
        {
            return new Token(val, TokenType, IsNegated);
        }

        /// <summary>
        /// Clones the token with a new value set for isNegated.
        /// </summary>
        /// <param name="isNegated">The new isNegated value</param>
        /// <returns>A cloned Token</returns>
        internal Token CloneWithNegatedValue(bool isNegated)
        {
            if (
                (TokenType & TokenType.Decimal) == 0
                ||
                (TokenType & TokenType.Integer) == 0
                ||
                (TokenType & TokenType.ExcelAddress) == 0)
            {
                return new Token(Value, TokenType, isNegated);
            }
            return this;
        }

        /// <summary>
        /// Overrides object.ToString()
        /// </summary>
        /// <returns>TokenType, followed by value</returns>
        public override string ToString()
        {
            return TokenType.ToString() + ", " + Value;
        }
    }
}
