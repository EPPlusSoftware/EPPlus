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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    internal class TokenFactory : ITokenFactory
    {
        public TokenFactory(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1=false)
            : this(new TokenSeparatorProvider(), nameValueProvider, functionRepository, r1c1)
        {

        }

        public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider, IFunctionNameProvider functionNameProvider, bool r1c1)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
            _functionNameProvider = functionNameProvider;
            _nameValueProvider = nameValueProvider;
            _r1c1 = r1c1;
        }

        private readonly ITokenSeparatorProvider _tokenSeparatorProvider;
        private readonly IFunctionNameProvider _functionNameProvider;
        private readonly INameValueProvider _nameValueProvider;
        private bool _r1c1;
        public Token Create(IEnumerable<Token> tokens, string token)
        {
            return Create(tokens, token, null);
        }
        public Token Create(IEnumerable<Token> tokens, string token, string worksheet)
        {
            Token tokenSeparator = default(Token);
            if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out tokenSeparator))
            {
                return tokenSeparator;
            }
            var tokenList = tokens.ToList();
            //Address with worksheet-string before  /JK
            if (token.StartsWith("!", StringComparison.OrdinalIgnoreCase) && tokenList[tokenList.Count - 1].TokenTypeIsSet(TokenType.String))
            {
                string addr = "";
                var i = tokenList.Count - 2;
                if (i > 0)
                {
                    if (tokenList[i].TokenTypeIsSet(TokenType.StringContent))
                    {
                        addr = "'" + tokenList[i].Value.Replace("'", "''") + "'";
                    }
                    else
                    {
                        throw (new ArgumentException(string.Format("Invalid formula token sequence near {0}", token)));
                    }
                    //Remove the string tokens and content
                    tokenList.RemoveAt(tokenList.Count - 1);
                    tokenList.RemoveAt(tokenList.Count - 1);
                    tokenList.RemoveAt(tokenList.Count - 1);

                    return new Token(addr + token, TokenType.ExcelAddress);
                }
                else
                {
                    throw (new ArgumentException(string.Format("Invalid formula token sequence near {0}", token)));
                }

            }
            if (string.IsNullOrEmpty(worksheet) && tokenList.Count>0)
            {
                if (tokenList[tokenList.Count - 1].TokenTypeIsSet(TokenType.WorksheetNameContent))
                {
                    worksheet = tokenList[tokenList.Count - 1].Value;
                }
                else if(tokenList.Count > 1 && tokenList[tokenList.Count - 1].TokenTypeIsSet(TokenType.WorksheetName) && tokenList[tokenList.Count - 2].TokenTypeIsSet(TokenType.WorksheetNameContent))
                {
                    worksheet = tokenList[tokenList.Count - 2].Value;
                }
            }
            if (tokens.Any() && tokens.Last().TokenTypeIsSet(TokenType.String))
            {
                return new Token(token, TokenType.StringContent);
            }
            if (!string.IsNullOrEmpty(token))
            {
                token = token.Trim();
            }
            if (IsNumeric(token, false))
            {
                return new Token(token, TokenType.Integer);
            }
            if (IsNumeric(token, true))
            {
                return new Token(token, TokenType.Decimal);
            }
            if (token.Equals("true", StringComparison.InvariantCultureIgnoreCase) || token.Equals("false", StringComparison.InvariantCultureIgnoreCase))
            {
                return new Token(token, TokenType.Boolean);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture).Contains("#REF!"))
            {
                return new Token(token, TokenType.InvalidReference);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#NUM!")
            {
                return new Token(token, TokenType.NumericError);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#VALUE!")
            {
                return new Token(token, TokenType.ValueDataTypeError);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#NULL!")
            {
                return new Token(token, TokenType.Null);
            }
            if (_nameValueProvider != null && _nameValueProvider.IsNamedValue(token, worksheet))
            {
                return new Token(token, TokenType.NameValue);
            }
            if (_functionNameProvider.IsFunctionName(token))
            {
                return new Token(token, TokenType.Function);
            }
            if (tokenList.Count > 0 && tokenList[tokenList.Count - 1].TokenTypeIsSet(TokenType.OpeningEnumerable))
            {
                return new Token(token, TokenType.Enumerable);
            }
            var at = OfficeOpenXml.ExcelAddressBase.IsValid(token, _r1c1);
            if (at==ExcelAddressBase.AddressType.InternalAddress || at == ExcelAddressBase.AddressType.ExternalAddress)
            {
                return new Token(token, TokenType.ExcelAddress);
            } 
            else if (at == ExcelAddressBase.AddressType.R1C1)
            {
                return new Token(token, TokenType.ExcelAddressR1C1);
            }
            else if(at==ExcelAddressBase.AddressType.InternalName || at == ExcelAddressBase.AddressType.ExternalName)
            {
                return new Token(token, TokenType.NameValue);
            }
            return new Token(token, TokenType.Unrecognized);

        }

        public static bool IsNumeric(string value, bool allowDecimal)
        {
            var nExp = 0;
            var nDot = 0;
            var nMinus = 0;
            var nPlus = 0;
            foreach(var c in value)
            {
                if ((c < '0' || c > '9') && (allowDecimal == false || c != '.') && c != 'E' && (c != '-' && c != '+'))
                {
                    return false;
                }
                if (c == 'E') nExp++;
                if (c == '.') nDot++;
                if (c == '-') nMinus++;
                if (c == '+') nPlus ++;
            }
            if(nExp == 0 && nMinus == 0)
            {
                return true;
            }
            else if(nExp == 1 && nDot == 1 && (nMinus == 1 || nPlus == 1))
            {
                return double.TryParse(value, NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out double r);
            }
            return false;
        }

        public Token Create(string token, TokenType explicitTokenType)
        {
            return new Token(token, explicitTokenType);
        }
    }
}
