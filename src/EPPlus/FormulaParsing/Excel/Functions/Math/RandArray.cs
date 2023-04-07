/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "7",
        Description = "Returns an array of random numbers",
        SupportsArrays = true)]
    internal class RandArray : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var nRows = 1;
            short nCols = 1;
            var min = 0d;
            var max = 1d;
            bool useInteger = false;
            if (arguments != null && arguments.Any())
            {
                if (arguments.Count() > 0)
                {
                    nRows = ArgToInt(arguments, 0);
                }
                if (arguments.Count() > 1)
                {
                    var c = ArgToInt(arguments, 1);
                    if (c > short.MaxValue || c < short.MinValue)
                    {
                        return CreateResult(eErrorType.Value);
                    }
                    nCols = Convert.ToInt16(c);
                }
                if (arguments.Count() > 2)
                {
                    min = ArgToDecimal(arguments, 2);
                }
                if (arguments.Count() > 3)
                {
                    max = ArgToDecimal(arguments, 3);
                }
                if (arguments.Count() > 4)
                {
                    useInteger = ArgToBool(arguments, 4);
                }
            }
            // 50 million cells in the array is the max value
            if(nRows * nCols > 50000000)
            {
                return CreateResult(eErrorType.Value);
            }
            else if(max < min)
            {
                return CreateResult(eErrorType.Value);
            }
            var rnd = new Random();
            var result = new InMemoryRange(new RangeDefinition(nRows, nCols));
            for (var row = 0; row < nRows; row++)
            {
                for (short col = 0; col < nCols; col++)
                {
                    var num = GetRandomNumber(rnd, min, max, useInteger);
                    result.SetValue(row, col, num);
                }
            }
            return CreateResult(result, DataType.ExcelRange);
        }

        private double GetRandomNumber(Random rnd, double min, double max, bool useInteger)
        {
            if(!useInteger)
            {
                var randomNumber = rnd.NextDouble();
                var span = System.Math.Abs(max - min);
                return min + span * randomNumber;
            }
            else
            {
                return rnd.Next((int)min, (int)max);
            }
            
        }
    }
}
