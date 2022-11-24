using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class OptimizedSourceCodeTokenizer : ISourceCodeTokenizer
    {
        private static readonly Dictionary<char, Token> _charAddressTokens = new Dictionary<char, Token>
        {
            {'!', new Token("!", TokenType.ExcelAddress)},
            {'$', new Token("$", TokenType.ExcelAddress)},
        };
        private static readonly Dictionary<char, Token> _charTokens = new Dictionary<char, Token>
        {
            {'+', new Token("+",TokenType.Operator)},
            {'-', new Token("-", TokenType.Operator)},
            {'*', new Token("*", TokenType.Operator)},
            {'/', new Token("/", TokenType.Operator)},
            {'^', new Token("^", TokenType.Operator)},
            {'&', new Token("&", TokenType.Operator)},
            {'>', new Token(">", TokenType.Operator)},
            {'<', new Token("<", TokenType.Operator)},
            {'=', new Token("=", TokenType.Operator)},
            {':', new Token(":", TokenType.Operator)},
            {'(', new Token("(", TokenType.OpeningParenthesis)},
            {')', new Token(")", TokenType.ClosingParenthesis)},
            {'{', new Token("{", TokenType.OpeningEnumerable)},
            {'}', new Token("}", TokenType.ClosingEnumerable)},
            {'\"', new Token("\"", TokenType.String)},
            {',', new Token(",", TokenType.Comma)},
            {';', new Token(";", TokenType.SemiColon) },
            {'%', new Token("%", TokenType.Percent) },
            {' ', new Token(" ", TokenType.WhiteSpace) },
            {'[', new Token("[", TokenType.OpeningBracket)},
            {']', new Token("]", TokenType.ClosingBracket) },
            {'!', new Token("!", TokenType.WorksheetName) }
        };
        private static readonly Dictionary<string, Token> _stringTokens = new Dictionary<string, Token>
        {
            {">=", new Token(">=", TokenType.Operator)},
            {"<=", new Token("<=", TokenType.Operator)},
            {"<>", new Token("<>", TokenType.Operator)},
        };
        private static readonly HashSet<string> _tableParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            {"#all" },
            {"#this row"},
            {"#headers" },
            {"#data" },
            {"#totals" }
        };
        private bool _r1c1, _keepWhitespace;
        /// <summary>
        /// The default tokenizer. This tokenizer will remove and ignore whitespaces.
        /// </summary>
        public static ISourceCodeTokenizer Default
        {
            get { return new OptimizedSourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false, false); }
        }
        /// <summary>
        /// The tokenizer used for r1c1 format. This tokenizer will keep whitespaces and add them as tokens.
        /// </summary>
        public static ISourceCodeTokenizer R1C1
        {
            get { return new OptimizedSourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, true, true); }
        }

        public OptimizedSourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false, bool keepWhitespace=false)
        {
            _r1c1 = r1c1;
            _keepWhitespace = keepWhitespace;
        }
        public OptimizedSourceCodeTokenizer(ITokenFactory tokenFactory)
        {
            _tokenFactory = tokenFactory;
        }

        private readonly ITokenFactory _tokenFactory;

        /// <summary>
        /// Split the input string into tokens used by the formula parser
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public IList<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        [Flags]
        enum statFlags : int
        {
           isString =   0x1,
           isOperator = 0x2,
           isAddress =  0x4,
           isNonNumeric=0x8,
           isNumeric = 0x10,
           isDecimal = 0x20,
           isPercent = 0x40,
           isNegator = 0x80,
           isColon =   0x100,
           isTableRef= 0x200,
           isExtRef =  0x400,
           isIntersect= 0x800,
        }
        public IList<Token> Tokenize(string input, string worksheet)
        {
            var l = new List<Token>();
            int ix;
            var length = input.Length;
            
            if (length > 0 && (input[0] == '+' /*|| input[0] == '='*/))
            {
                ix=1;
            }
            else
            {
                ix = 0;
            }

            statFlags flags = 0;
            short isInString = 0;
            short bracketCount = 0, paranthesesCount=0;
            var current =new StringBuilder();
            var pc = '\0';
            var separatorTokens = TokenSeparatorProvider.Instance.Tokens;
            while (ix < length)
            {
                var c = input[ix];
                if(c == '\"' && isInString != 2)
                {
                    if (pc == c && isInString == 0)
                    {
                        current.Append(c);
                    }
                    else
                    {
                        flags |= statFlags.isString;
                    }
                    isInString ^= 1;
                }
                else if (c == '\'' && isInString !=1)
                {
                    current.Append(c);
                    isInString ^= 2;
                }
                else
                { 
                    if(isInString==0 && _charTokens.ContainsKey(c))
                    {
                        if(c=='!' && current.Length > 0 && current[0]=='#')
                        {
                            var currentString=current.ToString(); 
                            if (currentString.Equals("#NUM", StringComparison.OrdinalIgnoreCase))
                            {
                                l.Add(new Token("#NUM!", TokenType.NumericError));
                            }
                            else if (currentString.Equals("#VALUE", StringComparison.OrdinalIgnoreCase))
                            {
                                l.Add(new Token("#VALUE!", TokenType.ValueDataTypeError));
                            }
                            else if (currentString.Equals("#NULL", StringComparison.OrdinalIgnoreCase))
                            {
                                l.Add(new Token("#NULL!", TokenType.Null));
                            }
                            else
                            {
                                l.Add(new Token("#REF!", TokenType.InvalidReference));
                            }
                            flags &= statFlags.isTableRef;
#if (NET35)
                            current=new StringBuilder();
#else
                            current.Clear();
#endif
                        }
                        else if (c==' ' && bracketCount > 0)
                        {
                            current.Append(c);
                        }
                        else
                        {
                            HandleToken(l, c, current, ref flags);

                            if (c == '-')
                            {
                                flags |= statFlags.isNegator;
                            }
                            else if (c == '[')
                            {
                                if ((flags & statFlags.isTableRef) != statFlags.isTableRef)
                                {
                                    flags |= statFlags.isExtRef;
                                }
                                l.Add(_charTokens[c]);
                            }
                            else if (c=='+' && l.Count>0) //remove leading + and add + operator.
                            {
                                var pt = GetLastToken(l);

                                //Remove prefixing +
                                if (!(pt.TokenType==TokenType.Operator
                                    ||
                                    pt.TokenType==TokenType.Negator
                                    ||
                                    pt.TokenType==TokenType.OpeningParenthesis
                                    ||
                                    pt.TokenType==TokenType.Comma
                                    ||
                                    pt.TokenType==TokenType.SemiColon
                                    ||
                                    pt.TokenType==TokenType.OpeningEnumerable))
                                {
                                    l.Add(_charTokens[c]);
                                }
                            }
                            else if (ix + 1 < length && _stringTokens.ContainsKey(input.Substring(ix, 2)))
                            {
                                l.Add(_stringTokens[input.Substring(ix, 2)]);
                                ix++;
                            }
                            else
                            {
                                if(c==' ')
                                {
                                    short wsCnt = 1;
                                    int wsIx = ix + 1;
                                    while(wsIx < input.Length && input[wsIx++]==' ')
                                    {
                                        wsCnt++;
                                    }
                                    var pt = GetLastToken(l);
                                    if(pt.TokenType==TokenType.CellAddress || 
                                       pt.TokenType == TokenType.ClosingParenthesis ||
                                       pt.TokenType == TokenType.NameValue ||
                                       pt.TokenType == TokenType.InvalidReference)
                                    {
                                        flags |= statFlags.isIntersect;
                                    }

                                    if (_keepWhitespace)
                                    {
                                        l.Add(new Token(new string(c, wsCnt), TokenType.WhiteSpace));
                                    }
                                    ix = wsIx >= input.Length && input[input.Length - 1] == ' ' ? wsIx - 1 : wsIx - 2;
                                }
                                else
                                {
                                    l.Add(_charTokens[c]);
                                }
                            }

                            if (c=='(')
                            {
                                paranthesesCount++;
                            }
                            else if(c==')')
                            {
                                paranthesesCount--;
                            }
                            else if (c == '[')
                            {
                                bracketCount++;
                            }
                            else if (c == ']')
                            {
                                bracketCount--;
                                if(bracketCount==0)
                                {
                                    flags = 0;
                                }
                            }
                        }
                    }
                    else
                    {
                        if(isInString==0)
                        {
                            if (current.Length == 0 && c == ':' && pc == ')')
                            {
                                l.Add(new Token(":", TokenType.Colon));
                                SetRangeOffsetToken(l);
                                flags |= statFlags.isColon;
                            }
                            else
                            {
                                if (_charAddressTokens.ContainsKey(c)) //handel :
                                {
                                    flags |= statFlags.isAddress;
                                }
                                else if (c >= '0' && c <= '9')
                                {
                                    flags |= statFlags.isNumeric;
                                }
                                else if (c == '.')
                                {
                                    flags |= statFlags.isDecimal;
                                }
                                else if (c == '%')
                                {
                                    flags |= statFlags.isPercent;
                                }
                                else
                                {
                                    flags |= statFlags.isNonNumeric;
                                }
                                current.Append(c);
                            }
                        }
                        else
                        {
                            current.Append(c);
                        }
                    }
                }
                ix++;
                if(c != ' ') pc = c;
            }
            if (current.Length > 0 || (flags & statFlags.isString) == statFlags.isString)
            {
                HandleToken(l, pc, current, ref flags);
            }

            if (isInString != 0)
            {
                throw new FormatException("Unterminated string");
            }
            else if (paranthesesCount != 0)
            {
                throw new FormatException("Number of opened and closed parentheses does not match");
            }
            else if (bracketCount != 0)
            {
                throw new FormatException("Number of opened and closed brackets does not match");
            }

            return l;
        }

        private Token GetLastToken(List<Token> l)
        {
            var i = l.Count - 1;
            while (i >= 0 && l[i].TokenType == TokenType.WhiteSpace) 
                i--;
            return l[i];
        }

        private void SetRangeOffsetToken(List<Token> l)
        {
            int i = l.Count - 1;
            int p= 0;
            while (i >= 0)
            {
                if ((l[i].TokenType & TokenType.OpeningParenthesis) == TokenType.OpeningParenthesis)
                {
                    p--;
                }
                else if((l[i].TokenType & TokenType.ClosingParenthesis) == TokenType.ClosingParenthesis)
                {
                    p++;
                }
                else if ((l[i].TokenType & TokenType.Function) == TokenType.Function && l[i].Value.Equals("offset", StringComparison.OrdinalIgnoreCase) && p==0)
                {
                    l[i] = new Token(l[i].Value, TokenType.RangeOffset | TokenType.Function);
                }
                i--;
            }
        }

        private void HandleToken(List<Token> l,char c, StringBuilder current, ref statFlags flags)
        {
            if ((flags & statFlags.isNegator) == statFlags.isNegator)
            {
                if (l.Count == 0)
                {
                    l.Add(new Token("-", TokenType.Negator));
                }
                else
                {
                    var pt = GetLastToken(l);
                    if((pt.TokenTypeIsSet(TokenType.Operator) && pt.Value == "+"))  //Replace + by -
                    {
                        l[l.Count - 1] = _charTokens['-'];
                    }
                    else if (pt.TokenType==TokenType.Operator
                        ||
                        pt.TokenType == TokenType.Negator
                        ||
                        pt.TokenType == TokenType.OpeningParenthesis
                        ||
                        pt.TokenType == TokenType.Comma
                        ||
                     pt.TokenType == TokenType.SemiColon
                        ||
                        pt.TokenType == TokenType.OpeningEnumerable)
                    {
                        l.Add(new Token("-", TokenType.Negator));
                    }
                    else
                    {
                        l.Add(_charTokens['-']);
                    }
                }
            }
            if (current.Length == 0)
            {
                if ((flags & statFlags.isIntersect) == statFlags.isIntersect)
                {                    
                    if (c == '[' ||
                        c == '(')
                    {
                        l.Add(new Token("isc", TokenType.Operator));
                    }
                }

                if ((flags & statFlags.isString) == statFlags.isString)
                {
                    l.Add(new Token("", TokenType.StringContent));
                }
                
                flags &= statFlags.isTableRef;
                return;
            }
            var currentString = current.ToString();

            if ((flags & statFlags.isString) == statFlags.isString)
            {
                l.Add(new Token(currentString, TokenType.StringContent));
            }
            else if (c == '!')
            {
                l.Add(new Token(currentString, TokenType.WorksheetNameContent));
            }
            else if (c == ']')
            {
                if(l.Count==0 || (flags & statFlags.isExtRef)== statFlags.isExtRef)
                {
                    l.Add(new Token(currentString, TokenType.ExternalReference));
                }
                else
                {
                    if((flags & statFlags.isTableRef) == statFlags.isTableRef)
                    { 
                        if(_tableParts.Contains(currentString))
                        {
                            l.Add(new Token(currentString, TokenType.TablePart));
                        }
                        else
                        {
                            l.Add(new Token(currentString, TokenType.TableColumn));
                        }
                    }
                }
            }
            else if (c == '(')
            {
                if((flags & statFlags.isColon) == statFlags.isColon)
                {
                    l.Add(new Token(currentString, TokenType.Function | TokenType.RangeOffset));
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.Function));
                }
            }            
            else if ((flags & statFlags.isAddress) == statFlags.isAddress)
            {
                if (_r1c1 == true)
                {
                    if (ExcelAddressBase.IsR1C1(currentString))
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddressR1C1));
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.NameValue));
                    }
                }
                else
                {
                    if (IsName(currentString))
                    {
                        l.Add(new Token(currentString, TokenType.NameValue));
                    }
                    else
                    {
                        //l.Add(new Token(currentString, TokenType.ExcelAddress));
                        l.Add(new Token(currentString, TokenType.CellAddress));
                    }
                }
            }
            else if ((flags & statFlags.isNonNumeric) == statFlags.isNonNumeric)
            {
                if (currentString.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   currentString.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.Boolean));
                }
                else if (_r1c1==false && IsValidCellAddress(currentString))
                {
                    l.Add(new Token(currentString, TokenType.CellAddress));
                }
                else if (_r1c1 == true && ExcelAddressBase.IsR1C1(currentString))
                {
                    if ((flags & statFlags.isColon) == statFlags.isColon)
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddressR1C1 | TokenType.RangeOffset));
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddressR1C1));
                    }
                }
                else
                {
                    if(c=='[')
                    {
                        l.Add(new Token(currentString, TokenType.TableName));
                        flags |= statFlags.isTableRef;
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.NameValue));
                    }
                }
            }
            else
            {
                if ((flags & statFlags.isPercent) == statFlags.isPercent)
                {
                    l.Add(new Token(currentString, TokenType.Percent));
                }
                else if ((flags & statFlags.isDecimal) == statFlags.isDecimal)
                {
                    l.Add(new Token(currentString, TokenType.Decimal));
                }
                else if ((flags & statFlags.isNumeric) == statFlags.isNumeric)
                {
                    l.Add(new Token(currentString, TokenType.Integer));
                }
            }

            if ((flags & statFlags.isIntersect) == statFlags.isIntersect)
            {
                var pt = GetLastToken(l);
                if (pt.TokenType == TokenType.CellAddress ||
                   pt.TokenType == TokenType.NameValue ||
                   pt.TokenType == TokenType.Function)
                {
                    l.Insert(l.Count - 1, new Token("isc", TokenType.Operator));
                }
            }

            flags &= statFlags.isTableRef;
            
            //Clear sb
#if(NET35)
    current=new StringBuilder();
#else
    current.Clear();
#endif
        }
    private static readonly char[] _addressChars = new char[]{':','$', '[', ']', '\''};
    private static bool IsName(string s)
    {
        var ix = s.LastIndexOf('!');
        if(ix>=0)
        {
            s = s.Substring(ix + 1);
        }        
        if (s.IndexOfAny(_addressChars) >=0) return false;
        return IsValidCellAddress(s)==false;
    }

        private static bool IsValidCellAddress(string address)
        {
            var numPos = -1;
            for (var i=0; i < address.Length; i++)
            {
                var c = address[i];
                if (c>='0' && c<='9')
                {
                    if (i == 0) return false;
                    if(numPos == -1) numPos = i;
                }
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    if (numPos != -1 || i > 3) return false;
                }
                else
                {
                    return false;
                }
            }
            if (numPos < 1 || numPos > 3) return false;
            var col = ExcelAddressBase.GetColumnNumber(address.Substring(0,numPos));
            if (col <= 0 || col > ExcelPackage.MaxColumns) return false;
            if(int.TryParse(address.Substring(numPos), out int row))
            {
                return row > 0 && row <= ExcelPackage.MaxRows; 

            }
            return false;
        }
    }
}
