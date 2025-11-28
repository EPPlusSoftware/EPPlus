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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "8.2",
        Description = "Applies a LAMBDA to each column and returns an array of the results. For example, if the original array is 3 columns by 2 rows, the returned array is 3 columns by 1 row.",
        IntroducedInExcelVersion = "2021")]
    internal class ByCol : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override string NamespacePrefix => "_xlfn.";

        public override bool ExecutesLambda => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = ArgToRangeInfo(arguments, 0);
            if (arguments[1].DataType != DataType.LambdaCalculation || arguments[1].Value is not LambdaCalculator calculator || calculator.NumberOfVariables != 1)
            {
                return CreateDynamicArrayResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
            }
            var resultRange = new InMemoryRange(1, range.Size.NumberOfCols);
            for(var col = 0; col < range.Size.NumberOfCols; col++)
            {
                var colRange = range.GetOffset(0, col, range.Size.NumberOfRows - 1, col);
                calculator.BeginCalculation();
                calculator.SetVariableValue(0, colRange, DataType.ExcelRange, context);
                var result = calculator.Execute(context);
                resultRange.SetValue(0, col, result.ResultValue);
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
        }
    }
}
