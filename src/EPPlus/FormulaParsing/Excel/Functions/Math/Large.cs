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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the Kth LARGEST value from a list of supplied numbers, for a given value K")]
    internal class Large : HiddenValuesHandlingFunction
    {
        public Large()
        {
            IgnoreHiddenValues = false;
            IgnoreErrors = false;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0);
            var index = ArgToInt(arguments, 1, IgnoreErrors) - 1;
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> {args}, context);
            if (index < 0 || index >= values.Count()) return CreateResult(eErrorType.Num);
            var result = values.OrderByDescending(x => x).ElementAt(index);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
