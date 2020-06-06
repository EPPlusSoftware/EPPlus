/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "4",
        Description = "Converts a non-number value to a number, a date to a serial number, the logical value TRUE to 1 and all other values to 0")]
    internal class N : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = GetFirstValue(arguments);
            
            if (arg is bool)
            {
                var val = (bool) arg ? 1d : 0d;
                return CreateResult(val, DataType.Decimal);
            }
            else if (IsNumeric(arg))
            {
                var val = ConvertUtil.GetValueDouble(arg);
                return CreateResult(val, DataType.Decimal);
            }
            else if (arg is string)
            {
                return CreateResult(0d, DataType.Decimal);
            }
            else if (arg is ExcelErrorValue)
            {
                return CreateResult(arg, DataType.ExcelError);
            }
            throw new ExcelErrorValueException(eErrorType.Value);
        }
    }
}
