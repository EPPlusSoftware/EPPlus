/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaFunctionExpression : VariableFunctionExpression
    {
        internal override bool IsLambda => true;
        private readonly List<LambdaStoredExpression> _storedExpressions = new List<LambdaStoredExpression>();
        
        internal LambdaFunctionExpression(string tokenValue, Stack<FunctionExpression> funcStack, ParsingContext ctx, int pos) : base(tokenValue, funcStack, ctx, pos)
        {
        }

        public List<LambdaStoredExpression> StoredExpressions => _storedExpressions;


        internal void AddExpression(Expression exp1, Expression exp2, IOperator op)
        {
            var storedExpression = new LambdaStoredExpression(exp1, exp2, op);
            _storedExpressions.Add(storedExpression);
        }

        public override CompileResult Compile()
        {
            var calculator = new LambdaCalculator(StoredExpressions, _args.ToList());
            return new CompileResult(calculator, DataType.LambdaCalculation);
        }

        internal class LambdaStoredExpression
        {
            private readonly Stack<Expression> _expressions = new Stack<Expression>();

            private readonly IOperator _operator;

            public Stack<Expression> Expressions => _expressions;

            public IOperator Operator => _operator;

            public LambdaStoredExpression(Expression exp1, Expression exp2, IOperator op)
            {
                _expressions.Push(exp2);
                _expressions.Push(exp1);
                _operator = op;
            }
        }
    }
}
