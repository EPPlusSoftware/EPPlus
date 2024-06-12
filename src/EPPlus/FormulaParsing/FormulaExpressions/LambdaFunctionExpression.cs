﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaFunctionExpression : VariableFunctionExpression
    {
        internal override bool IsLambda => true;
        private List<Token> _lambdaTokens;
        
        internal LambdaFunctionExpression(string tokenValue, Stack<FunctionExpression> funcStack, ParsingContext ctx, int pos) : base(tokenValue, funcStack, ctx, pos)
        {

        }
    }
}
