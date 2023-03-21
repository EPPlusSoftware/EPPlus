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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "4",
        Description = "Returns information about the data type of a supplied value")]
    internal class TypeFunction : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var val = arguments.ElementAt(0).Value;
            if (val is bool)
                return CreateResult(4, DataType.Integer);
            if (IsNumeric(val) || val == null)
                return CreateResult(1, DataType.Integer);
            if (ExcelErrorValue.Values.IsErrorValue(val))
                return CreateResult(16, DataType.Integer);
            if (val is string)
                return CreateResult(2, DataType.Integer);
            if (val.GetType().IsArray || val is IEnumerable)
                return CreateResult(64, DataType.Integer);
            return new CompileResult(eErrorType.Value);
        }
    }
}
