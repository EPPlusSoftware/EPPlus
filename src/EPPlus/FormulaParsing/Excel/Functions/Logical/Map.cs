using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    internal class Map : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

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
            var ranges = new List<IRangeInfo>();
            for (var i = 0; i < arguments.Count - 1; i++) 
            {
                var rng = ArgToRangeInfo(arguments, i);
                ranges.Add(rng);
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
                var rng = ranges[i];
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
            for (var row = 0; row < maxRows; row++)
            {
                for (var col = 0; col < maxCols; col++)
                {
                    calculator.BeginCalculation();
                    for(var rIx = 0; rIx < ranges.Count; rIx++)
                    {
                        var rng = ranges[rIx];
                        if(rng.Size.NumberOfCols > col)
                        {
                            var value = rng.GetOffset(row, col);
                            var cr = CompileResultFactory.Create(value);
                            calculator.SetVariableValue(rIx, value, cr.DataType, context);
                        }
                        else
                        {
                            calculator.SetVariableValue(rIx, ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError, context);
                        }
                        
                    }
                    var compileResult = calculator.Execute(context);
                    resultRange.SetValue(row, col, compileResult.ResultValue);
                }
            }
            return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
        }
    }
}
