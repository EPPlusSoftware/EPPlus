/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/12/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class IsOmittedExpression : VariableFunctionExpression
    {
        internal IsOmittedExpression(string tokenValue, ParsingContext ctx, int pos) 
            : base(tokenValue, ctx, pos, false)
        {
        }

        internal override void AddArgument(int arg)
        {
            base.AddArgument(arg);
        }

        internal override bool HandlesVariables => true;

        internal override bool IsVariableArg(int arg, bool isLastArgument)
        {
            return true;
        }
    }
}
