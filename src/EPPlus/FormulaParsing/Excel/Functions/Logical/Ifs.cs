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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "5.0",
        Description = "Returns the largest numeric value that meets one or more criteria in a range of values",
        IntroducedInExcelVersion = "2019")]
    internal class Ifs : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var maxArgs = arguments.Count() < (127 * 2) ? arguments.Count() : 127 * 2; 
            for(var x = 0; x < maxArgs; x += 2)
            {
                if (System.Math.Round(ArgToDecimal(arguments, x), 15) != 0d) return CompileResultFactory.Create(arguments.ElementAt(x + 1).Value);
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
        }
        public override bool ReturnsReference => true;
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if (argumentIndex % 2 == 0)
            {
                return FunctionParameterInformation.Condition;
            }
            else
            {
                return FunctionParameterInformation.UseIfConditionIsTrue;
            }
        }
    }
}
