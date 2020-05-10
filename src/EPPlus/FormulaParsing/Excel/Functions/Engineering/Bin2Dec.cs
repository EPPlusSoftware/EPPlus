/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Helpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class Bin2Dec : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToString(arguments, 0);
            if (number.Length > 10) return CreateResult(eErrorType.Num);
            if (number.Length < 10)
            {
                return CreateResult(Convert.ToInt32(number, 2), DataType.Decimal);
            }
            if (!BinaryHelper.TryParseBinaryToDecimal(number, 2, out int result)) return CreateResult(eErrorType.Num);
            return CreateResult(result, DataType.Integer);
        }
    }
}
