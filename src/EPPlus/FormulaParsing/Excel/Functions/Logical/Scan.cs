/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/04/2025        EPPlus Software AB       Initial release EPPlus 8.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "8.1",
        Description = "Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value.",
        IntroducedInExcelVersion = "2021")]
    internal class Scan : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var argIx = 0;
            object initialValue = null;
            if (arguments[0].IsExcelRange)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
            }
            initialValue = arguments[0].Value;
            var ivDataType = arguments[0].DataType;
            if (initialValue is ExcelErrorValue e) return CreateResult(e.Type);
            var range = ArgToRangeInfo(arguments, 1);
            if (arguments[2].Value is not LambdaCalculator calculator) return CreateResult(eErrorType.Value);
            var accumulatedValue = initialValue;
            var resultRange = new InMemoryRange(range.Size);
            for (var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var rangeValue = range.GetOffset(row, col);
                    var cr = CompileResultFactory.Create(rangeValue);
                    calculator.BeginCalculation();
                    if(ivDataType == DataType.Empty) ivDataType = cr.DataType;
                    if(ivDataType != DataType.String && accumulatedValue == null)
                    {
                        accumulatedValue = 0;
                    }
                    calculator.SetVariableValue(0, accumulatedValue, ivDataType, context);
                    calculator.SetVariableValue(1, rangeValue, cr.DataType, context);
                    var compileResult = calculator.Execute(context);
                    accumulatedValue = compileResult.Result;
                    resultRange.SetValue(row, col, accumulatedValue);
                }
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
        }
    }
}
