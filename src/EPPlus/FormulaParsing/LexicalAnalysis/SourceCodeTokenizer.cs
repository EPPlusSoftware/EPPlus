using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Source code tokenizer
    /// </summary>
    public class SourceCodeTokenizer : ISourceCodeTokenizer
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
            {'!', new Token("!", TokenType.WorksheetName) },
            {'\'',  new Token("\'", TokenType.SingleQuote) },
            //{'#',  new Token("#'", TokenType.HashMark) },
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
        private bool _r1c1, _keepWhitespace, _isPivotFormula;
        private readonly TokenType _nameValueOrPivotFieldToken;
		/// <summary>
		/// The default tokenizer. This tokenizer will remove and ignore whitespaces.
		/// </summary>
		public static ISourceCodeTokenizer Default
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false, false, false); }
        }

        /// <summary>
        /// The tokenizer used for r1c1 format. This tokenizer will keep whitespaces and add them as tokens.
        /// </summary>
        public static ISourceCodeTokenizer R1C1
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, true, true, false); }
        }
		/// <summary>
		/// The default tokenizer. This tokenizer will remove and ignore whitespaces.
		/// </summary>
		public static ISourceCodeTokenizer Default_KeepWhiteSpaces
		{
			get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false, true); }
		}

		/// <summary>
		/// </summary>
		public static ISourceCodeTokenizer PivotFormula
		{
			get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false, false, true); }
		}
		/// <param name="functionRepository">A function name provider</param>
		/// <param name="nameValueProvider">A name value provider</param>
		/// <param name="r1c1">If true the tokenizer will use the R1C1 format</param>
		/// <param name="keepWhitespace">If true whitspaces in formulas will be preserved</param>
        /// <param name="pivotFormula">If the formula is from a calculated column in a pivot table.</param>
		public SourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false, bool keepWhitespace = false, bool pivotFormula = false)
		{
			_r1c1 = r1c1;
            _keepWhitespace = keepWhitespace;
            _isPivotFormula = pivotFormula;
            _nameValueOrPivotFieldToken = _isPivotFormula ? TokenType.PivotField : TokenType.NameValue;
		}
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
            isString = 0x1,
            isOperator = 0x2,
            isAddress = 0x4,
            isNonNumeric = 0x8,
            isNumeric = 0x10,
            isDecimal = 0x20,
            isPercent = 0x40,
            isNegator = 0x80,
            isColon = 0x100,
            isTableRef = 0x200,
            isExtRef = 0x400,
            isIntersect = 0x800,
            isError = 0x1000,
            isExponential = 0x2000,
            isLastCharQuote = 0x4000
        }
        /// <summary>
        /// Splits a formula in tokens used in when calculating for example a worksheet cell, defined name or a pivot table field formula.
        /// </summary>
        /// <param name="input">The formula to tokenize</param>
        /// <param name="worksheet">The worksheet name.</param>
        /// <returns></returns>
        /// <exception cref="InvalidFormulaException">Thrown if the formula is not valid.</exception>
         public IList<Token> Tokenize(string input, string worksheet)
        {
            var l = new List<Token>();
            int ix;
            var length = input.Length;

            if (length > 0 && (input[0] == '+' /*|| input[0] == '='*/))
            {
                ix = 1;
            }
            else
            {
                ix = 0;
            }

            statFlags flags = 0;
            short isInString = 0;
            short bracketCount = 0, paranthesesCount = 0;
            var current = new StringBuilder();
            var pc = '\0';
            var isR1C1 = false;
            List<int> variableFuncPositions = default;
            while (ix < length)
            {
                var c = input[ix];
                if (c == '\"' && isInString != 2 && bracketCount==0)
                {
                    current.Append(c);
                    flags |= statFlags.isString;
                    isInString ^= 1;
                }
                else if (c == '\'' && isInString != 1)
                {
                    if (bracketCount == 0)
                    {
                        if (_isPivotFormula)
                        {
                            if(current.Length == 0)
                            {
                                flags |= statFlags.isNonNumeric;
                            }
							else if (pc == '\'')
                            {
								current.Append(c);
								flags |= statFlags.isLastCharQuote;
							}
						}
                        else
                        {
							if (current.Length == 0)
							{
								l.Add(_charTokens['\'']);
							}
							else
							{
								current.Append(c);
							}
						}
						isInString ^= 2;
                    }
                    else if (pc == '\'' && (flags & statFlags.isLastCharQuote)==0)
                    {
                        current.Append(c);
                        flags |= statFlags.isLastCharQuote;
                    }
                    else
                    {
                        flags &= ~statFlags.isLastCharQuote;
						current.Append(c);
					}
				}
                else
                {
                    if (bracketCount == 0 && isInString == 0 && IsWhiteSpace(c))
                    {
                        HandleToken(l, c, ref current, ref flags, ref variableFuncPositions);
                        short wsCnt = 1;
                        int wsIx = ix + 1;
                        while (wsIx < input.Length && IsWhiteSpace(input[wsIx++]))
                        {
                            wsCnt++;

                        }
                        if ((flags & statFlags.isNegator) != statFlags.isNegator)
                        {
                            var pt = GetLastToken(l);
                            if (pt.TokenType == TokenType.CellAddress ||
                                pt.TokenType == TokenType.ClosingParenthesis ||
                                pt.TokenType == TokenType.NameValue ||
                                pt.TokenType == TokenType.InvalidReference)
                            {
                                flags |= statFlags.isIntersect;
                            }
                        }

                        if (_keepWhitespace)
                        {
                            l.Add(new Token(input.Substring(ix, wsCnt), TokenType.WhiteSpace));
                        }
                        ix = wsIx >= input.Length && IsWhiteSpace(input[input.Length - 1]) ? wsIx - 1 : wsIx - 2;
                    }
                    else if (isInString == 0 && _charTokens.ContainsKey(c) && (flags & statFlags.isExponential) == 0)
                    {
                        
                        if (c == '!' && current.Length > 0 && current[0] == '#')
                        {
                            var currentString = current.ToString();
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
                            current = new StringBuilder();
                        }
                        else if (c == '/' && current.Length == 2 && current[0] == '#' && (current[1] == 'n' || current[1] == 'N'))
                        {
                            current.Append(c); //We have a #n/a
                        }
                        else if ((((c != '[' && c != ']' && c != '\'')) || ((c == '[' || c == ']' || c == '\'') && (pc == '\'' && (flags & statFlags.isLastCharQuote) == 0))) && !(c == ',' && pc == ']' && current.Length == 0) && bracketCount > 0)
                        {
                            current.Append(c);
                        }
                        else if (_r1c1 && (c == '[' && (pc == 'R' || pc == 'C' || pc == 'r' || pc == 'c')) || ((c == '-' || c == ']' || c == ':' || (c >= '0' && c <= '9')) && isR1C1)) //Handel [-1]
                        {
                            isR1C1 = c != ']';
                            current.Append(c);
                            if (ix == input.Length - 1)
                            {
                                if ((flags & statFlags.isNegator) == statFlags.isNegator)
                                {
                                    HandleNegator(l, current, flags);
                                }
                                l.Add(new Token(current.ToString(), TokenType.ExcelAddressR1C1));
                                return l;
                            }
                        }
                        else
                        {
                            HandleToken(l, c, ref current, ref flags, ref variableFuncPositions);

                            if (c == '-')
                            {
                                flags |= statFlags.isNegator;
                            }
                            else if (c == '[')
                            {
                                if (bracketCount > 0 && pc == '\'')
                                {
                                    current.Append(c);
                                }
                                else
                                {
                                    if ((flags & statFlags.isTableRef) != statFlags.isTableRef)
                                    {
                                        flags |= statFlags.isExtRef;
                                    }
                                    l.Add(_charTokens[c]);
                                }
                            }
                            else if (c == '+' && l.Count > 0) //remove leading + and add + operator.
                            {
                                var pt = GetLastToken(l);

                                //Remove prefixing +
                                if (!(pt.TokenType == TokenType.Operator
                                    ||
                                    pt.TokenType == TokenType.Negator
                                    ||
                                    pt.TokenType == TokenType.OpeningParenthesis
                                    ||
                                    pt.TokenType == TokenType.Comma
                                    ||
                                    pt.TokenType == TokenType.SemiColon
                                    ||
                                    pt.TokenType == TokenType.OpeningEnumerable))
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
                                l.Add(_charTokens[c]);
                            }

                            if (c == '(')
                            {
                                paranthesesCount++;
                            }
                            else if (c == ')')
                            {
                                paranthesesCount--;
                            }
                            else if (c == '[' && (pc != '\'' || (flags & statFlags.isLastCharQuote) == 0))
                            {
                                bracketCount++;
                            }
                            else if (c == ']' && (pc != '\'' || (flags & statFlags.isLastCharQuote) == 0))
                            {
                                bracketCount--;
                                if (bracketCount == 0)
                                {
                                    flags = 0;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (isInString == 0)
                        {
                            if (current.Length == 0 && c == ':' && pc == ')')
                            {
                                l.Add(new Token(":", TokenType.Colon));
                                SetRangeOffsetToken(l);
                                flags |= statFlags.isColon;
                            }
                            else if ((flags == statFlags.isNumeric || 
                                      flags == (statFlags.isNumeric | statFlags.isDecimal) || 
                                      flags == (statFlags.isNumeric | statFlags.isDecimal | statFlags.isNegator) || 
                                      flags == (statFlags.isNumeric | statFlags.isNegator)) 
                                      && 
                                      (flags & statFlags.isNonNumeric) != statFlags.isNonNumeric 
                                      && 
                                      (c == 'E' || c == 'e')) //Handle exponential values in a formula.
                            {
                                current.Append(c);
                                flags |= statFlags.isExponential;
                            }
                            else if ((flags & statFlags.isExponential) == statFlags.isExponential && (c == '+' || c == '-'))
                            {
                                current.Append(c);
                                flags &= ~statFlags.isExponential;
                                flags |= statFlags.isDecimal;
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
                if (c != ' ') pc = c;
            }
            if (current.Length > 0 || (flags & statFlags.isString) == statFlags.isString)
            {
                HandleToken(l, pc, ref current, ref flags, ref variableFuncPositions);
            }
            if (isInString != 0)
            {
                throw new InvalidFormulaException("Unterminated string");
            }
            else if (paranthesesCount != 0)
            {
                throw new InvalidFormulaException("Number of opened and closed parentheses does not match");
            }
            else if (bracketCount != 0)
            {
                throw new InvalidFormulaException("Number of opened and closed brackets does not match");
            }

            if(variableFuncPositions != null && variableFuncPositions.Count > 0)
            {
                var variableHelper = new VariableParameterHelper(l, variableFuncPositions);
                variableHelper.Process();
            }
            return l;
        }
#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        private bool IsWhiteSpace(char c)
        {
            return char.IsWhiteSpace(c) || c == '\r' || c == '\n' || c == '\t';
        }

        private Token GetLastToken(List<Token> l)
        {
            if (l.Count == 0) return default(Token);
            var i = l.Count - 1;
            while (i > 0 && l[i].TokenType == TokenType.WhiteSpace)
                i--;
            return l[i];
        }
        private Token GetLastTokenIgnore(List<Token> l, out int i, params TokenType[] ignoreTokens)
        {
            i = l.Count - 1;
            while (i > 0 && ignoreTokens.Contains(l[i].TokenType))
                i--;
            return l[i];
        }


        private void SetRangeOffsetToken(List<Token> l)
        {
            int i = l.Count - 1;
            int p = 0;
            while (i >= 0)
            {
                if ((l[i].TokenType & TokenType.OpeningParenthesis) == TokenType.OpeningParenthesis)
                {
                    p--;
                }
                else if ((l[i].TokenType & TokenType.ClosingParenthesis) == TokenType.ClosingParenthesis)

                {
                    p++;
                }
                else if ((l[i].TokenType & TokenType.Function) == TokenType.Function && l[i].Value.Equals("offset", StringComparison.OrdinalIgnoreCase) && p == 0)
                {
                    l[i] = new Token(l[i].Value, TokenType.RangeOffset | TokenType.Function);
                }
                i--;
            }
        }

        private bool IsParameterVariable(string token)
        {
            return token.StartsWith("_xlpm.");
        }
        private void HandleToken(List<Token> l, char c, ref StringBuilder current, ref statFlags flags, ref List<int> variableFuncPositions)
        {
            if ((flags & statFlags.isNegator) == statFlags.isNegator)
            {
                HandleNegator(l, current, flags);
            }
            if (current.Length == 0)
            {
                if ((flags & statFlags.isIntersect) == statFlags.isIntersect)
                {
                    if (c == '[' ||
                        c == '(')
                    {
                        l.Add(new Token(Operator.IntersectIndicator, TokenType.Operator));
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
                if (GetLastToken(l).TokenType == TokenType.SingleQuote)
                {
                    var ix = currentString.IndexOf(']', 1);
                    if (ix > 1)
                    {
                        var extId = currentString.Substring(1, ix - 1);
                        l.Add(_charTokens['[']);
                        l.Add(new Token(extId, TokenType.ExternalReference));
                        l.Add(_charTokens[']']);
                        ix++;
                    }
                    else
                    {
                        ix = 0;
                    }
                    l.Add(new Token(currentString.Substring(ix, currentString.Length - ix - 1), TokenType.WorksheetNameContent));
                    l.Add(_charTokens['\'']);
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.WorksheetNameContent));
                }
            }
            else if (c == ']')
            {
                if (l.Count == 0 || (flags & statFlags.isExtRef) == statFlags.isExtRef)
                {
                    l.Add(new Token(currentString, TokenType.ExternalReference));
                }
                else
                {
                    if ((flags & statFlags.isTableRef) == statFlags.isTableRef)
                    {
                        if (_tableParts.Contains(currentString))
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
                if (VariableParameterHelper.IsVariableParameterFunction(currentString))
                {
                    if(variableFuncPositions == default)
                    {
                        variableFuncPositions = new List<int>();
                    }
                    variableFuncPositions.Add(l.Count);
                }
                if ((flags & statFlags.isColon) == statFlags.isColon)
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
                        l.Add(new Token(currentString, _nameValueOrPivotFieldToken));
                    }
                }
                else
                {
                    if (IsName(currentString))
                    {
                        l.Add(new Token(currentString, _nameValueOrPivotFieldToken));
                    }
                    else if(IsParameterVariable(currentString))
                    {
                        l.Add(new Token(currentString, TokenType.ParameterVariable));
                    }
                    else
                    {

                        if ((flags & statFlags.isNumeric) == statFlags.isNumeric && (flags & statFlags.isNonNumeric) == statFlags.isNonNumeric)
                        {
                            l.Add(new Token(currentString, TokenType.CellAddress));
                        }
                        else if ((flags & statFlags.isNumeric) == statFlags.isNumeric)
                        {
                            l.Add(new Token(currentString, TokenType.FullRowAddress));
                        }
                        else
                        {
                            l.Add(new Token(currentString, TokenType.FullColumnAddress));
                        }
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
                else if (currentString.Equals("#N/A", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.NAError));
                }
                else if (_r1c1 == false && IsValidCellAddress(currentString))
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
                    if (c == '[')
                    {
                        l.Add(new Token(currentString, TokenType.TableName));
                        flags |= statFlags.isTableRef;
                    }
                    else
                    {
                        if ((c == ':' || (l.Count > 0 && l[l.Count - 1] == _charTokens[':'])) && ExcelCellBase.IsColumnLetter(currentString))   //We have a full column address
                        {
                            l.Add(new Token(currentString, TokenType.FullColumnAddress));
                        }
                        else
                        {
                            l.Add(new Token(currentString, _nameValueOrPivotFieldToken));
                        }
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
                    if ((c == ':' || (l.Count > 0 && l[l.Count - 1] == _charTokens[':'])))   //We have a full row address
                    {
                        l.Add(new Token(currentString, TokenType.FullRowAddress));
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.Integer));
                    }
                }
                else if ((flags & statFlags.isTableRef) == statFlags.isTableRef && currentString == ":")
                {
                    l.Add(_charTokens[':']);
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.InvalidReference));
                }
            }

            if ((flags & statFlags.isIntersect) == statFlags.isIntersect)
            {
                var pt = GetLastToken(l);
                if (pt.TokenType == TokenType.CellAddress ||
                   pt.TokenType == TokenType.NameValue ||
                   pt.TokenType == TokenType.Function)
                {
                    if (_keepWhitespace && l.Count > 2 && l[l.Count - 2].TokenType==TokenType.WhiteSpace)
                    {
                        var wsToken = l[l.Count - 2];
						if (wsToken.Value.Length > 1) //Multiple white space?
                        {
                            wsToken.Value = wsToken.Value.Substring(0, wsToken.Value.Length);
							l.Insert(l.Count - 1, new Token(Operator.IntersectIndicator, TokenType.Operator));
						}
                        else
                        {
							l[l.Count - 2] = new Token(Operator.IntersectIndicator, TokenType.Operator);
						}
					}
					else
                    {
						l.Insert(l.Count - 1, new Token(Operator.IntersectIndicator, TokenType.Operator));
					}
				}
            }

            flags &= statFlags.isTableRef;

            //Clear sb
            current = new StringBuilder();
        }

        private void HandleNegator(List<Token> l, StringBuilder current, statFlags flags)
        {
            if (l.Count == 0)
            {
                if ((flags & statFlags.isNonNumeric) == 0 && (flags & statFlags.isNumeric) == statFlags.isNumeric)
                {
                    current.Insert(0, '-');
                }
                else
                {
                    l.Add(new Token("-", TokenType.Negator));
                }
            }
            else
            {
                var pt = GetLastTokenIgnore(l, out int index, TokenType.SingleQuote, TokenType.WorksheetNameContent, TokenType.ExternalReference, TokenType.OpeningBracket, TokenType.WhiteSpace);
                if (pt.TokenType == TokenType.Operator
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
                    if ((flags & statFlags.isNonNumeric) == 0 && (flags & statFlags.isNumeric) == statFlags.isNumeric)
                    {
                        current.Insert(0, '-');
                    }
                    else
                    {
                        InsertNegatorToken(l, pt, index, new Token("-", TokenType.Negator));
                    }
                }
                else
                {
                    InsertNegatorToken(l, pt, index, _charTokens['-']);
                }
            }
        }

        private static void InsertNegatorToken(List<Token> l, Token pt, int index, Token token)
        {

            if (pt != l[l.Count - 1])
            {
                if (l[index + 1].TokenType == TokenType.WhiteSpace)
                {
                    l.Insert(index + 2, token);
                }
                else
                {
                    l.Insert(index + 1, token);
                }
            }
            else
            {
                l.Add(token);
            }
        }

        private static readonly char[] _addressChars = new char[] { ':', '$', '[', ']', '\'' };
        private static bool IsName(string s)
        {
            var ix = s.LastIndexOf('!');
            if (ix >= 0)
            {
                s = s.Substring(ix + 1);
            }
            if (s.IndexOfAny(_addressChars) >= 0) return false;
            return IsValidCellAddress(s) == false;
        }

        private static bool IsValidCellAddress(string address)
        {
            var numPos = -1;
            for (var i = 0; i < address.Length; i++)
            {
                var c = address[i];
                if (c >= '0' && c <= '9')
                {
                    if (i == 0) return false;
                    if (numPos == -1) numPos = i;
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
            int col;
            //if (numPos == -1) //Column reference only, for exampel A:A
            //{
            //    col = ExcelAddressBase.GetColumnNumber(address);
            //    return col > 0 && col <= ExcelPackage.MaxColumns;
            //}
            if (numPos < 1 || numPos > 3)
            {
                return false;
            }
            col = ExcelCellBase.GetColumn(address.Substring(0, numPos));
            if (col <= 0 || col > ExcelPackage.MaxColumns) return false;
            if (int.TryParse(address.Substring(numPos), out int row))
            {
                return row > 0 && row <= ExcelPackage.MaxRows;

            }
            return false;
        }
    }
}
