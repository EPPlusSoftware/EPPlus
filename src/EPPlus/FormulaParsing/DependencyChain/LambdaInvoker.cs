/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/03/2025         EPPlus Software AB       EPPlus 8.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal static class LambdaInvoker
    {
        // This method compiles the arguments for a Lambda function,
        // invokes it and returns the result.
        internal static CompileResult InvokeLambdaFunction(RpnOptimizedDependencyChain depChain, RpnFormula f)
        {
            var lambdaArgs = new List<CompileResult>();
            if (!f._expressionStack.Any(x => x.ExpressionType == ExpressionType.LambdaCalculation)) return null;
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
                    else if(exp != null)
                    {
                        result = exp.Compile();
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
