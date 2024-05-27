using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculator
    {
        public LambdaCalculator(List<LambdaFunctionExpression.LambdaStoredExpression> expressions, List<CompileResult> variables)
        {
            CalculationExpressions = expressions;
            Variables = variables;
        }

        public List<LambdaFunctionExpression.LambdaStoredExpression> CalculationExpressions { get; private set; }

        public List<CompileResult> Variables { get; private set; }

        public void SetVariableValue(int index, object value)
        {
            var variable = Variables[index];
            var exps = new List<VariableExpression>();
            foreach(var storedExpression in CalculationExpressions)
            {
                var e1 = storedExpression.Expressions.Pop();
                var e2 = storedExpression.Expressions.Pop();
                storedExpression.Expressions.Push(e2);
                storedExpression.Expressions.Push(e1);
                if(e1 is VariableExpression ve1 && ve1.Name == variable.Result.ToString())
                {
                    ve1.SetValue(ve1.Name, CompileResultFactory.Create(value));
                }
                if (e2 is VariableExpression ve2 && ve2.Name == variable.Result.ToString())
                {
                    ve2.SetValue(ve2.Name, CompileResultFactory.Create(value));
                }
            }
        }

        public CompileResult Execute(ParsingContext ctx)
        {
            foreach(var storedExpression in CalculationExpressions)
            {
                var e1 = storedExpression.Expressions.Pop();
                var e2 = storedExpression.Expressions.Pop();
                storedExpression.Expressions.Push(e2);
                storedExpression.Expressions.Push(e1);
                return storedExpression.Operator.Apply(e1.Compile(), e2.Compile(), ctx);
            }
            return CompileResult.Empty;
        }
    }
}
