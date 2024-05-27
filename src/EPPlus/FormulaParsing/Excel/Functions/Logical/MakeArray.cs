/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "7.2",
        Description = "Returns a calculated array of a specified row and column size, by applying a LAMBDA",
        IntroducedInExcelVersion = "2021")]
    internal class MakeArray : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rows = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if(e1 != null)
            {
                return CompileResult.GetErrorResult(e1.Type);
            }
            var cols = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null)
            {
                return CompileResult.GetErrorResult(e2.Type);
            }
            var arg3 = arguments[2];
            if(arg3.DataType != DataType.LambdaCalculation)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var lambdaCalculator = arg3.Value as LambdaCalculator;
            if (lambdaCalculator == null) return CompileResult.GetErrorResult(eErrorType.Value);
            var resultRange = new InMemoryRange(rows, (short)cols);
            for(var row = 0; row < rows; row++)
            {
                for(var col = 0; col  < cols; col++)
                {
                    lambdaCalculator.SetVariableValue(0, row + 1);
                    lambdaCalculator.SetVariableValue(1, col + 1);
                    var result = lambdaCalculator.Execute(context);
                    resultRange.SetValue(row, col, result.ResultValue);
                }
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
        }
    }
}
