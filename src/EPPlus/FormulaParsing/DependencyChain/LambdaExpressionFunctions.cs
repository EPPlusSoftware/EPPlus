using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal static class LambdaExpressionFunctions
    {
        internal static bool LastExpressionIsLambdaCalculation(Stack<Expression> s, out LambdaCalculationExpression lce)
        {
            lce = null;
            if (s.Count > 0 && s.Peek().ExpressionType == ExpressionType.LambdaCalculation && s.Peek() is LambdaCalculationExpression lce2)
            {
                lce = lce2;
                return true;
            }
            return false;
        }

        internal static void PreProcessLambdaCalculation(Stack<Expression> s, RpnFormula f, LambdaCalculationExpression lce)
        {

            var lambdaSettings = f.LambdaSettings;
            if (lambdaSettings.CurrentLambdaExpressions.Count == 0 || lambdaSettings.CurrentLambdaExpressions.Peek().Expression.Id != lce.Id)
            {
                lce.BeginCalculation();
                var stackPos = new LambdaExpressionStackPosition(s.Count - 1, lce);
                lambdaSettings.CurrentLambdaExpressions.Push(stackPos);
                lambdaSettings.LambdaArgsAdded.Push(0);
            }
        }

        internal static bool CheckLambdaExpression(Stack<Expression> s, RpnFormula f, LambdaExpressionStackPosition stackPos, ParsingContext ctx, out CompileResult compileResult)
        {
            compileResult = null;
            if (stackPos != null && (f.LambdaSettings.LambdaArgsAdded.Count > 0 ? f.LambdaSettings.LambdaArgsAdded.Peek() : 0) >= stackPos.Expression.GetNumberOfVariables())
            {
                var cr = stackPos.Expression.Compile();
                var calculator = cr.Result as LambdaCalculator;
                compileResult = calculator.Execute(ctx);
                if (compileResult.ResultType == CompileResultType.DynamicArray_AlwaysSetCellAsDynamic)
                {
                    f._flags |= FormulaFlags.IsAlwaysDynamic;
                }

                s.Pop();
                f.LambdaSettings.CurrentLambdaExpressions.Pop();
                f.LambdaSettings.LambdaArgsAdded.Pop();
                return true;
            }
            return false;
        }
    }
}
