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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        private int _tokenIndex = 0;
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

        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            _tokenIndex = 0;
            _graph.Reset();
            var tokensArr = tokens != null ? tokens.ToArray() : new Token[0];
            BuildUp(tokensArr, null);
            return _graph;
        }

        private void BuildUp(Token[] tokens, Expression parent)
        {
            while (_tokenIndex < tokens.Length)
            {
                var token = tokens[_tokenIndex];
                IOperator op = null;
                if (token.TokenTypeIsSet(TokenType.Operator) && OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    SetOperatorOnExpression(parent, op);
                }
                else if (token.TokenTypeIsSet(TokenType.Function))
                {
                    BuildFunctionExpression(tokens, parent, token.Value);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    _tokenIndex++;
                    BuildEnumerableExpression(tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    _tokenIndex++;
                    BuildGroupExpression(tokens, parent);
                    //if (parent is FunctionExpression)
                    //{
                    //    return;
                    //}
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingParenthesis) || token.TokenTypeIsSet(TokenType.ClosingEnumerable))
                {
                    break;
                }
                else if(token.TokenTypeIsSet(TokenType.WorksheetName))
                {
                    var sb = new StringBuilder();
                    sb.Append(tokens[_tokenIndex++].Value);
                    sb.Append(tokens[_tokenIndex++].Value);
                    sb.Append(tokens[_tokenIndex++].Value);
                    sb.Append(tokens[_tokenIndex].Value);
                    var t = new Token(sb.ToString(), TokenType.ExcelAddress);
                    CreateAndAppendExpression(ref parent, ref t);
                }
                else if (token.TokenTypeIsSet(TokenType.Negator))
                {
                    _negateNextExpression = !_negateNextExpression;
                }
                else if(token.TokenTypeIsSet(TokenType.Percent))
                {
                    SetOperatorOnExpression(parent, Operator.Percent);
                    if (parent == null)
                    {
                        _graph.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else
                {
                    CreateAndAppendExpression(ref parent, ref token);
                }
                _tokenIndex++;
            }
        }

        private void BuildEnumerableExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new EnumerableExpression());
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                var enumerableExpression = new EnumerableExpression();
                parent.AddChild(enumerableExpression);
                BuildUp(tokens, enumerableExpression);
            }
        }

        private void CreateAndAppendExpression(ref Expression parent, ref Token token)
        {
            if (IsWaste(token)) return;
            if (parent != null && 
                (token.TokenTypeIsSet(TokenType.Comma) || token.TokenTypeIsSet(TokenType.SemiColon)))
            {
                parent = parent.PrepareForNextChild();
                return;
            }
            if (_negateNextExpression)
            {
                token = token.CloneWithNegatedValue(true);
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token);
            if (parent == null)
            {
                _graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenTypeIsSet(TokenType.String))
            {
                return true;
            }
            return false;
        }

        private void BuildFunctionExpression(Token[] tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                _graph.Add(new FunctionExpression(funcName, _parsingContext, _negateNextExpression));
                _negateNextExpression = false;
                HandleFunctionArguments(tokens, _graph.Current);
            }
            else
            {
                var func = new FunctionExpression(funcName, _parsingContext, _negateNextExpression);
                _negateNextExpression = false;
                parent.AddChild(func);
                HandleFunctionArguments(tokens, func);
            }
        }

        private void HandleFunctionArguments(Token[] tokens, Expression function)
        {
            _tokenIndex++;
            var token = tokens.ElementAt(_tokenIndex);
            if (!token.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                throw new ExcelErrorValueException(eErrorType.Value);
            }
            _tokenIndex++;
            BuildUp(tokens, function.Children.First());
        }

        private void BuildGroupExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new GroupExpression(_negateNextExpression));
                _negateNextExpression = false;
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                if (parent.IsGroupedExpression || parent is FunctionArgumentExpression)
                {
                    var newGroupExpression = new GroupExpression(_negateNextExpression);
                    _negateNextExpression = false;
                    parent.AddChild(newGroupExpression);
                    BuildUp(tokens, newGroupExpression);
                }
                 BuildUp(tokens, parent);
            }
        }

        private void SetOperatorOnExpression(Expression parent, IOperator op)
        {
            if (parent == null)
            {
                _graph.Current.Operator = op;
            }
            else
            {
                Expression candidate;
                if (parent is FunctionArgumentExpression)
                {
                    candidate = parent.Children.Last();
                }
                else
                {
                    candidate = parent.Children.Last();
                    if (candidate is FunctionArgumentExpression)
                    {
                        candidate = candidate.Children.Last();
                    }
                }
                candidate.Operator = op;
            }
        }
    }
}
