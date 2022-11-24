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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal class RpnFormulaCell
    {
        public RpnFormulaCell()
        {
            _expressionStack = new Stack<RpnExpression>();
            _funcStackPosition = new Stack<int>();

        }
        internal Stack<RpnExpression> _expressionStack;
        internal Stack<int> _funcStackPosition;

    }
}