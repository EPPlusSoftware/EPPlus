using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal static class LambdaInvoker
    {
        // This method compiles the arguments for a Lambda function,
        // invokes it and returns the result.
        internal static CompileResult InvokeLambdaFunction(RpnOptimizedDependencyChain depChain, RpnFormula f)
        {
            var lambdaArgs = new List<CompileResult>();
            CompileResult result = default;
            while (f._expressionStack.Count > 0)
            {
                var exp = f._expressionStack.Pop();

                if (f._expressionStack.Count > 0 && exp.ExpressionType != ExpressionType.LambdaCalculation)
                {
                    var arg = exp.Compile();
                    lambdaArgs.Insert(0, arg);
                }
                else
                {
                    // The last expression on the stack should be a LambdaCalculationExpression
                    if (exp is LambdaCalculationExpression lce)
                    {
                        var lie = new LambdaInvokeExpression(lce, depChain._parsingContext, f._tokenIndex);
                        foreach (var arg in lambdaArgs)
                        {
                            lie.AddArgument(arg);
                        }
                        result = lie.Compile();
                        break;
                    }
                    else
                    {
                        f._expressionStack.Push(new ErrorExpression(new CompileResult(eErrorType.Value), depChain._parsingContext));
                        break;
                    }
                }
            }
            return result;
        }
    }
}
