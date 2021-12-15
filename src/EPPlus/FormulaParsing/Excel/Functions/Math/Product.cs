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
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the product of a supplied list of numbers")]
    internal class Product : HiddenValuesHandlingFunction
    {
        public Product()
        {
            IgnoreErrors = false;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var args = arguments.ToList();
            if(!IgnoreErrors && arguments.Any(x => x.ValueIsExcelError))
            {
                return CreateResult(arguments.First(x => x.ValueIsExcelError).ValueAsExcelErrorValue.Type);
            }
            args.RemoveAll(x => ShouldIgnore(x));
            var result = 1d;
            var values = ArgsToObjectEnumerable(true, args, context);
            foreach (var obj in values.Where(x => x != null && IsNumeric(x)))
            {
                result *= Convert.ToDouble(obj);
            }
            return CreateResult(result, DataType.Decimal);
        }

    }
}
