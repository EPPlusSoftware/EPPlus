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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Engineering,
        EPPlusVersion = "5.1",
        Description = "Converts a decimal number to hexadecimal")]
    internal class Dec2Hex : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var number = ArgToInt(arguments, 0);
            var padding = default(int?);
            if (arguments.Count() > 1)
            {
                padding = ArgToInt(arguments, 1);
                if (padding.Value < 0 ^ padding.Value > 10) return CreateResult(eErrorType.Num);
            }
            var result = Convert.ToString(number, 16);
            if(!string.IsNullOrEmpty(result))
            {
                result = result.ToUpper();
            }
            if(number < 0)
            {
                result = PaddingHelper.EnsureLength(result, 10, "F");
            }
            else if(padding.HasValue)
            {
                result = PaddingHelper.EnsureLength(result, padding.Value, "0");
            }
            else
            {
                result = PaddingHelper.EnsureMinLength(result, 10);
            }
            return CreateResult(result, DataType.String);
        }
    }
}
