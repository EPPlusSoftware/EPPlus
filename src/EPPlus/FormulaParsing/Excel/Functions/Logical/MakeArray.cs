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
        Description = "Returns a calculated array of a specified row and column size, by applying a LAMBDA",
        IntroducedInExcelVersion = "2021")]
    internal class MakeArray : ExcelFunction
    {
        public override int ArgumentMinLength => 3;

        public override bool ExecutesLambda => true;

        public override string NamespacePrefix => "_xlfn.";

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
            if (lambdaCalculator.NumberOfVariables != 2) return CompileResult.GetErrorResult(eErrorType.Value);
            var resultRange = new InMemoryRange(rows, (short)cols);
            for(var row = 0; row < rows; row++)
            {
                for(var col = 0; col  < cols; col++)
                {
                    lambdaCalculator.BeginCalculation();
                    lambdaCalculator.SetVariableValue(0, row + 1, DataType.Integer, context);
                    lambdaCalculator.SetVariableValue(1, col + 1, DataType.Integer, context);
                    var result = lambdaCalculator.Execute(context);
                    resultRange.SetValue(row, col, result.ResultValue);
                }
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
        }
    }
}
