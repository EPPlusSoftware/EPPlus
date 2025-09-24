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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "8.2",
        Description = "Reduces an array to an accumulated value by applying a LAMBDA to each value and returning the total value in the accumulator. ",
        IntroducedInExcelVersion = "2021")]
    internal class Reduce : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 3;

        public override bool IsVolatile => true;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            object initialValue = null;
            var ivDataType = DataType.Unknown;
            if (arguments[0].DataType != DataType.Empty)
            {
                initialValue = arguments[0].Value;
                ivDataType = arguments[0].DataType;
            }
            if (initialValue is ExcelErrorValue e) return CreateResult(e.Type);

            // Last arg must be a Lambda expression
            var lastArg = arguments.LastOrDefault();
            if (lastArg == null)
            {
                return CreateResult(eErrorType.Value);
            }
            if (lastArg.DataType != DataType.LambdaCalculation)
            {
                return CreateResult(eErrorType.Value);
            }
            var calculator = lastArg.Value as LambdaCalculator;
            var accumulatedValue = initialValue;
            if (arguments[1].DataType == DataType.ExcelRange)
            {
                var range = ArgToRangeInfo(arguments, 1);
                if(calculator.IsEtaReducedLambdaFunction)
                {
                    HandleEtaReducedLambda(context, calculator, ref ivDataType, ref accumulatedValue, range);
                }
                else
                {
                    HandleNormalLambda(context, ref ivDataType, calculator, ref accumulatedValue, range);
                }
               
            }
            else
            {
                calculator.BeginCalculation();
                calculator.SetVariableValue(0, initialValue, ivDataType, context);
                calculator.SetVariableValue(1, arguments[1].Value, arguments[1].DataType, context);
                var compileResult = calculator.Execute(context);
                accumulatedValue = compileResult.Result;
            }
           

            var accResult = CompileResultFactory.Create(accumulatedValue);
            return CreateDynamicArrayResult(accumulatedValue, accResult.DataType, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
        }

        private static void HandleEtaReducedLambda(ParsingContext context, LambdaCalculator calculator, ref DataType ivDataType, ref object accumulatedValue, IRangeInfo range)
        {
            var function = calculator.EtaFunction;
            for (var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var rangeValue = range.GetOffset(row, col);
                    var arg0 = new FunctionArgument(CompileResultFactory.Create(accumulatedValue));
                    var arg1 = new FunctionArgument(CompileResultFactory.Create(rangeValue));
                    var cr = function.Execute(new List<FunctionArgument> { arg0, arg1 }, context);
                    accumulatedValue = cr.Result;
                    if (ivDataType == DataType.Unknown) ivDataType = cr.DataType;
                }
            }
        }

        private static void HandleNormalLambda(ParsingContext context, ref DataType ivDataType, LambdaCalculator calculator, ref object accumulatedValue, IRangeInfo range)
        {
            for (var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var rangeValue = range.GetOffset(row, col);
                    var cr = CompileResultFactory.Create(rangeValue);
                    calculator.BeginCalculation();
                    if (ivDataType == DataType.Unknown) ivDataType = cr.DataType;
                    if (ivDataType != DataType.String && accumulatedValue == null)
                    {
                        accumulatedValue = 0;
                    }
                    calculator.SetVariableValue(0, accumulatedValue, ivDataType, context);
                    calculator.SetVariableValue(1, rangeValue, cr.DataType, context);
                    var compileResult = calculator.Execute(context);
                    accumulatedValue = compileResult.Result;
                }
            }
        }
    }
}
