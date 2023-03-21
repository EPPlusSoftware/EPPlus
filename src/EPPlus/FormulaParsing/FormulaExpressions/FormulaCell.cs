/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2022         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class FormulaCell
    {
        public FormulaCell()
        {
            _expressionStack = new Stack<Expression>();
            _funcStackPosition = new Stack<FunctionExpression>();
        }
        internal Stack<Expression> _expressionStack;
        internal Stack<FunctionExpression> _funcStackPosition;

    }
}