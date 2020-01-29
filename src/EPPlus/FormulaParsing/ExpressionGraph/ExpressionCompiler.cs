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
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionCompiler : IExpressionCompiler
    {
        private IEnumerable<Expression> _expressions;
        private IExpressionConverter _expressionConverter;
        private ICompileStrategyFactory _compileStrategyFactory;

        public ExpressionCompiler()
            : this(new ExpressionConverter(), new CompileStrategyFactory())
        {
 
        }

        public ExpressionCompiler(IExpressionConverter expressionConverter, ICompileStrategyFactory compileStrategyFactory)
        {
            _expressionConverter = expressionConverter;
            _compileStrategyFactory = compileStrategyFactory;
        }

        public CompileResult Compile(IEnumerable<Expression> expressions)
        {
            _expressions = expressions;
            return PerformCompilation();
        }
        public CompileResult Compile(string worksheet, int row, int column, IEnumerable<Expression> expressions)
        {
            _expressions = expressions;
            return PerformCompilation(worksheet, row, column);
        }

        private CompileResult PerformCompilation(string worksheet="", int row=-1, int column=-1)
        {
            var compiledExpressions = HandleGroupedExpressions();
            while(compiledExpressions.Any(x => x.Operator != null))
            {
                var prec = FindLowestPrecedence();
                compiledExpressions = HandlePrecedenceLevel(prec);
            }
            if (_expressions.Any())
            {
                return compiledExpressions.First().Compile();
            }
            return CompileResult.Empty;
        }

        private IEnumerable<Expression> HandleGroupedExpressions()
        {
            if (!_expressions.Any()) return Enumerable.Empty<Expression>();
            var first = _expressions.First();
            var groupedExpressions = _expressions.Where(x => x.IsGroupedExpression);
            foreach(var groupedExpression in groupedExpressions)
            {
                var result = groupedExpression.Compile();
                if (result == CompileResult.Empty) continue;
                var newExp = _expressionConverter.FromCompileResult(result);
                newExp.Operator = groupedExpression.Operator;
                newExp.Prev = groupedExpression.Prev;
                newExp.Next = groupedExpression.Next;
                if (groupedExpression.Prev != null)
                {
                    groupedExpression.Prev.Next = newExp;
                }
                if (groupedExpression.Next != null)
                {
                    groupedExpression.Next.Prev = newExp;
                }
                if (groupedExpression == first)
                {
                    first = newExp;
                }
            }
            return RefreshList(first);
        }

        private IEnumerable<Expression> HandlePrecedenceLevel(int precedence)
        {
            var first = _expressions.First();
            var expressionsToHandle = _expressions.Where(x => x.Operator != null && x.Operator.Precedence == precedence);
            var last = expressionsToHandle.Last();
            var expression = expressionsToHandle.First();
            do
            {
                var strategy = _compileStrategyFactory.Create(expression);
                var compiledExpression = strategy.Compile();
                if(compiledExpression is ExcelErrorExpression)
                {
                    return RefreshList(compiledExpression);
                }
                if (expression == first)
                {
                    first = compiledExpression;
                }

                expression = compiledExpression;
            }
            while (expression != null && expression.Operator != null && expression.Operator.Precedence == precedence);
            return RefreshList(first);
        }

        private int FindLowestPrecedence()
        {
            return _expressions.Where(x => x.Operator != null).Min(x => x.Operator.Precedence);
        }

        private IEnumerable<Expression> RefreshList(Expression first)
        {
            var resultList = new List<Expression>();
            var exp = first;
            resultList.Add(exp);
            while (exp.Next != null)
            {
                resultList.Add(exp.Next);
                exp = exp.Next;
            }
            _expressions = resultList;
            return resultList;
        }
    }
}
