using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LetFunctionExpression : VariableFunctionExpression
    {
        internal LetFunctionExpression(string tokenValue, Stack<FunctionExpression> funcStack, ParsingContext ctx, int pos) : base(tokenValue, funcStack, ctx, pos)
        {

        }

        internal override void AddArgument(int arg)
        {
            base.AddArgument(arg);
        }

        internal override bool HandlesVariables => true;

        internal override bool IsVariableArg(int arg, bool isLastArgument)
        {
            return arg % 2 == 0 && !isLastArgument;
        }

    }
}
