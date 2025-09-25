/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/03/2025         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LetFunctionExpression : VariableFunctionExpression
    {
        internal LetFunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {

        }

        internal override bool IsLet => true;

        internal override void AddArgument(int arg)
        {
            base.AddArgument(arg);
        }

        internal override bool HandlesVariables => true;

        internal override bool IsVariableArg(int arg, bool isLastArgument)
        {
            return arg % 2 == 0 && !isLastArgument;
        }

        internal override void OnDispose()
        {
            base.OnDispose();
            if (Context.VariableStorage.Count > 0)
            {
                Context.VariableStorage.Pop();
            }

        }

    }
}
