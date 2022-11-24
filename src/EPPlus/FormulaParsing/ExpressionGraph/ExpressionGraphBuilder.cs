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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.ExternalReferences;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        //private ExpressionTree _graph = new ExpressionTree();
        private int _tokenIndex = 0;
        private FormulaAddressBase _currentAddress;
        private bool _negateNextExpression;
        public ExpressionGraphBuilder(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
        {

        }
        public ExpressionGraphBuilder(IExpressionFactory expressionFactory, ParsingContext parsingContext)
        {
            _expressionFactory = expressionFactory;
            _parsingContext = parsingContext;
        }
        public ExpressionTree Build(IEnumerable<Token> tokens)
        {
            _tokenIndex = 0;
            var graph=new ExpressionTree();
            var tokensArr = tokens != null ? tokens.ToArray() : new Token[0];
            BuildUp(graph, tokensArr, null);
            return graph;
        }

        private void BuildUp(ExpressionTree graph, Token[] tokens, Expression parent)
        {
            int bracketCount = 0;
            Expression rangeParent=null;
            while (_tokenIndex < tokens.Length)
            {
                var token = tokens[_tokenIndex];
                IOperator op = null;
                if (token.TokenTypeIsSet(TokenType.OpeningBracket))
                {
                    bracketCount++;
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingBracket))
                {
                    bracketCount--;
                    if(bracketCount==0 && _currentAddress is FormulaTableAddress ta)
                    {
                        ta.SetTableAddress(_parsingContext.Package);
                        CreateAndAppendExpression(graph, ref parent, ref token);
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.ExternalReference))
                {
                    _currentAddress = new FormulaAddressBase() { ExternalReferenceIx = (short)_parsingContext.Package.Workbook.ExternalLinks.GetExternalLink(token.Value) };
                }
                else if (token.TokenTypeIsSet(TokenType.WorksheetNameContent))
                {
                    if (_currentAddress == null)
                    {
                        _currentAddress = new FormulaAddressBase();
                    }
                    if(_currentAddress.ExternalReferenceIx == -1)
                    {
                        _currentAddress.WorksheetIx = _parsingContext.Package.Workbook.Worksheets.GetPositionByToken(token.Value);
                    }
                    else if(_currentAddress.ExternalReferenceIx > -1)
                    {
                        var er = _parsingContext.Package.Workbook.ExternalLinks[_currentAddress.ExternalReferenceIx];
                        if (er.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                        {
                            var wb = (ExcelExternalWorkbook)er;
                            if(wb.Package==null)
                            {
                                _currentAddress.WorksheetIx = (short)(wb.Package.Workbook.Worksheets[token.Value]?.SheetId ?? -1);
                            }
                            else 
                            {
                                _currentAddress.WorksheetIx = (short)(wb.CachedWorksheets[token.Value]?.SheetId ?? -1);
                            }
                        }
                        else
                        {
                            _currentAddress.WorksheetIx = -1;
                        }
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.TableName))
                {
                    if(_currentAddress!=null)
                    {
                        _currentAddress = new FormulaTableAddress()
                        {
                            ExternalReferenceIx = _currentAddress.ExternalReferenceIx,
                            WorksheetIx = _currentAddress.WorksheetIx, 
                            TableName = token.Value 
                        };
                    }
                    else
                    {
                        _currentAddress = new FormulaTableAddress() { TableName = token.Value };
                    }
                }
                else if(token.TokenTypeIsSet(TokenType.TableColumn))
                {
                    var ta = (FormulaTableAddress)_currentAddress;
                    if (string.IsNullOrEmpty(ta.ColumnName1))
                    {
                        ta.ColumnName1 = token.Value;
                    }
                    else
                    {
                        ta.ColumnName2 = token.Value;
                    }
                }
                else if(token.TokenTypeIsSet(TokenType.TablePart))
                {
                    var ta = (FormulaTableAddress)_currentAddress;
                    if (string.IsNullOrEmpty(ta.TablePart1))
                    {
                        ta.TablePart1 = token.Value;
                    }
                    else
                    {
                        ta.TablePart2 = token.Value;
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.Operator) && OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    if(!(bracketCount > 0 && op.Operator==Operators.Colon) && !(_tokenIndex==0 && op==Operator.Eq))
                    {
                        var current = GetCurrentExpression(graph, parent);

                        if ((op.Operator == Operators.Colon || op.Operator == Operators.Intersect))
                        {
                            if (!(parent is RangeExpression))
                            {
                                var rangeExpression = new RangeExpression(_parsingContext) { _parent = parent };
                                rangeExpression.Children.Add(current);
                                var exps = parent == null ? graph.Expressions : parent.Children;
                                exps.Remove(current);
                                exps.Add(rangeExpression);
                                rangeParent = parent;
                                ((ExpressionWithParent)parent)._parent = rangeExpression;
                                parent = rangeExpression;                                
                            }
                            current.Operator = op;
                        }
                        else if (rangeParent!=null)
                        {
                            parent.Operator = op;
                            parent = rangeParent;
                            rangeParent = null;
                        }
                        else
                        {
                            current.Operator = op;
                        }
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.Function))
                {                    
                    BuildFunctionExpression(graph, tokens, parent, token.Value);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    _tokenIndex++;
                    BuildEnumerableExpression(graph, tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    _tokenIndex++;
                    BuildGroupExpression(graph, tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingParenthesis) || token.TokenTypeIsSet(TokenType.ClosingEnumerable))
                {
                    break;
                }
                else if (token.TokenTypeIsSet(TokenType.Negator))
                {
                    _negateNextExpression = !_negateNextExpression;
                }
                else if(token.TokenTypeIsSet(TokenType.Percent))
                {
                    var current = GetCurrentExpression(graph, parent);
                    current.Operator = Operator.Percent;
                    if (parent == null)
                    {
                        graph.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else if(!
                    (token.TokenTypeIsSet(TokenType.Comma) && bracketCount > 0 ||
                     token.TokenTypeIsSet(TokenType.WhiteSpace) || 
                     token.TokenTypeIsSet(TokenType.WorksheetName))
                    )
                {
                    CreateAndAppendExpression(graph, ref parent, ref token);
                }
                _tokenIndex++;
            }
        }

        private void BuildEnumerableExpression(ExpressionTree graph, Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                graph.Add(new EnumerableExpression(_parsingContext));
                BuildUp(graph, tokens, graph.Current);
            }
            else
            {
                var enumerableExpression = new EnumerableExpression(_parsingContext);
                parent.AddChild(enumerableExpression);
                BuildUp(graph, tokens, enumerableExpression);
            }
        }

        private void CreateAndAppendExpression(ExpressionTree graph, ref Expression parent, ref Token token)
        {
            if (IsWaste(token)) return;
            if (parent != null && 
                (token.TokenTypeIsSet(TokenType.Comma) || token.TokenTypeIsSet(TokenType.SemiColon)))
            {
                parent = parent.PrepareForNextChild(token);
                if(parent is RangeExpression re)
                {
                    parent = re._parent;
                }
                return;
            }
            if (_negateNextExpression)
            {
                token = token.CloneWithNegatedValue(true);
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token, ref _currentAddress, parent);

            _currentAddress = null;
            if (parent == null)
            {
                graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenTypeIsSet(TokenType.String) || token.TokenTypeIsSet(TokenType.Colon))
            {
                return true;
            }
            return false;
        }

        //private void BuildRangeOffsetExpression(Token[] tokens, Expression parent, Token token, IDictionary<int, TokenInfo> tokenInfo)
        //{
        //    if(_nRangeOffsetTokens++ % 2 == 0)
        //    {
        //        _rangeOffsetExpression = new RangeOffsetExpression(_parsingContext);
        //        if(token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
        //        {
        //            _rangeOffsetExpression.OffsetExpression1 = new FunctionExpression("offset", _parsingContext, false);
        //            HandleFunctionArguments(tokens, _rangeOffsetExpression.OffsetExpression1, tokenInfo);
        //        }
        //        else if(token.TokenTypeIsSet(TokenType.ExcelAddress))
        //        {
        //            _rangeOffsetExpression.AddressExpression2 = _expressionFactory.Create(token) as ExcelAddressExpression;
        //        }
        //    }
        //    else
        //    {
        //        if (parent == null)
        //        {
        //            _graph.Add(_rangeOffsetExpression);
        //        }
        //        else
        //        {
        //            parent.AddChild(_rangeOffsetExpression);
        //        }
        //        if (token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
        //        {
        //            _rangeOffsetExpression.OffsetExpression2 = new FunctionExpression("offset", _parsingContext, _negateNextExpression);
        //            HandleFunctionArguments(tokens, _rangeOffsetExpression.OffsetExpression2, tokenInfo);
        //        }
        //        else if (token.TokenTypeIsSet(TokenType.ExcelAddress))
        //        {
        //            _rangeOffsetExpression.AddressExpression2 = _expressionFactory.Create(token) as ExcelAddressExpression;
        //        }
        //    }
        //}

        private void BuildFunctionExpression(ExpressionTree graph, Token[] tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                graph.Add(new FunctionExpression(funcName, _parsingContext, _negateNextExpression, parent));
                _negateNextExpression = false;
                HandleFunctionArguments(graph, tokens, graph.Current);
            }
            else
            {
                var func = new FunctionExpression(funcName, _parsingContext, _negateNextExpression, parent);
                _negateNextExpression = false;
                parent.AddChild(func);
                HandleFunctionArguments(graph, tokens, func);
            }
        }

        private void HandleFunctionArguments(ExpressionTree graph, Token[] tokens, Expression function)
        {
            _tokenIndex++;
            var token = tokens.ElementAt(_tokenIndex);
            if (!token.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                throw new ExcelErrorValueException(eErrorType.Value);
            }
            _tokenIndex++;
            //BuildUp(graph, tokens, function.Children.First());
            BuildUp(graph, tokens, function);
        }

        private void BuildGroupExpression(ExpressionTree graph, Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                graph.Add(new GroupExpression(_negateNextExpression, _parsingContext));
                _negateNextExpression = false;
                BuildUp(graph, tokens, graph.Current);
            }
            else
            {
                //if (parent.IsGroupedExpression || parent is FunctionArgumentExpression)
                if (parent.IsGroupedExpression || parent is FunctionExpression)
                {
                    var newGroupExpression = new GroupExpression(_negateNextExpression, _parsingContext);
                    _negateNextExpression = false;
                    parent.AddChild(newGroupExpression);
                    BuildUp(graph, tokens, newGroupExpression);
                }
                 BuildUp(graph, tokens, parent);
            }
        }
        private Expression GetCurrentExpression(ExpressionTree graph, Expression parent)
        {
            Expression current;
            if (parent == null)
            {
                current = graph.Current;
            }
            else
            {
                if (parent.ExpressionType==ExpressionType.Function)
                {
                    current = parent.Children.Last();
                }
                else
                {
                    current = parent.Children.Last();
                    if (current.ExpressionType==ExpressionType.Function)
                    {
                        current = current.Children.Last();
                    }
                }
            }

            return current;
        }
    }
}
