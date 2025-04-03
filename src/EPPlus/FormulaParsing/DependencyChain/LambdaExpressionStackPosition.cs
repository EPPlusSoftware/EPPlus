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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal class LambdaExpressionStackPosition
    {
        public LambdaExpressionStackPosition(int stackIndex, LambdaCalculationExpression exp)
        {
            StackIndex = stackIndex;
            Expression = exp;
        }
        public LambdaCalculationExpression Expression { get; private set; }

        public int StackIndex { get; set; }
    }
}
