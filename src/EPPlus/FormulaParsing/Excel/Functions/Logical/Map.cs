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
        Description = "Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA to create a new value.",
        IntroducedInExcelVersion = "2021")]
    internal class Map : ExcelFunction
    {
        private class RangeReference
        {
            public RangeReference(IRangeInfo rng)
            {
                Range = rng;
                Shift = 0;
            }

            public IRangeInfo Range { get; private set; }

            public int Shift { get; set; }

            public object GetByIndex(int index)
            {
                var nCols = Range.Size.NumberOfCols;
                var nRows = Range.Size.NumberOfRows;
                var ix = index + Shift;

                // Check if the index is outside the bounds of the matrix
                if (ix >= nCols * nRows)
                {
                    // Return the last element of the matrix
                    return Range.GetOffset(nRows - 1, nCols - 1);
                }

                var col = ix % nCols;
                var row = ix / nCols;
                return Range.GetOffset(row, col);
            }


        }

        public override int ArgumentMinLength => 2;

        public override string NamespacePrefix => "_xlfn.";

        public override bool ExecutesLambda => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var lastArg = arguments.LastOrDefault();
            if(lastArg == null)
            {
                return CreateResult(eErrorType.Value);
            }
            if(lastArg.DataType != DataType.LambdaCalculation)
            {
                return CreateResult(eErrorType.Value);
            }
            var calculator = lastArg.Value as LambdaCalculator;
            
            // get ranges from arguments
            var ranges = new List<RangeReference>();
            for (var i = 0; i < arguments.Count - 1; i++) 
            {
                var rng = ArgToRangeInfo(arguments, i);
                ranges.Add(new RangeReference(rng));
            }
            // there must be one variable per supplied range
            if(calculator.NumberOfVariables != ranges.Count)
            {
                return CreateResult(eErrorType.Value);
            }

            // find max rows and max cols (determine the size of the result range)
            var maxRows = 0;
            short maxCols = 0;
            for(var i = 0; i < ranges.Count; i++)
            {
                var rng = ranges[i].Range;
                if(rng.Size.NumberOfRows > maxRows)
                {
                    maxRows = rng.Size.NumberOfRows;
                }
                if(rng.Size.NumberOfCols > maxCols)
                {
                    maxCols = rng.Size.NumberOfCols;
                }
            };
            var resultRange = new InMemoryRange(maxRows, maxCols);

            // build result
            var ix = 0;
            for (var row = 0; row < maxRows; row++)
            {
                for (var col = 0; col < maxCols; col++)
                {
                    calculator.BeginCalculation();
                    for(var rIx = 0; rIx < ranges.Count; rIx++)
                    {
                        var rng = ranges[rIx];
                        if(rng.Range.Size.NumberOfCols > col && rng.Range.Size.NumberOfRows > row)
                        {
                            var value = rng.GetByIndex(ix);
                            var cr = CompileResultFactory.Create(value);
                            calculator.SetVariableValue(rIx, value, cr.DataType, context);
                        }
                        else
                        {
                            rng.Shift++;
                            calculator.SetVariableValue(rIx, ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError, context);
                        }
                        
                    }
                    ix++;
                    var compileResult = calculator.Execute(context);
                    resultRange.SetValue(row, col, compileResult.ResultValue);
                }
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange, CompileResultType.DynamicArray_AlwaysSetCellAsDynamic);
        }
    }
}
