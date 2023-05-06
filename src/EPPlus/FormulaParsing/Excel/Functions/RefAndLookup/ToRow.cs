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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Returns the array in a single row.",
        SupportsArrays = true)]
    internal class ToRow : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.First();
            var ignore = 0;
            if(arguments.Count() > 1 && arguments.ElementAt(1).Value != null) 
            {
                ignore = ArgToInt(arguments, 1);
                if(ignore < 0 || ignore > 4)
                {
                    return CompileResult.GetErrorResult(eErrorType.Value);
                }
            }
            var scanByColumn = false;
            if (arguments.Count() > 2 && arguments.ElementAt(2).Value != null)
            {
                scanByColumn = ArgToBool(arguments, 2);
            }
            if (firstArg.IsExcelRange)
            {
                var range = firstArg.ValueAsRangeInfo;
                var result = new List<object>();
                var maxX = scanByColumn ? range.Size.NumberOfCols : range.Size.NumberOfRows;
                var maxy = scanByColumn ? range.Size.NumberOfRows : range.Size.NumberOfCols;
                for (var x = 0; x < maxX; x++)
                {
                    for(short y = 0; y < maxy; y++)
                    {
                        var v = scanByColumn ? range.GetOffset(y, x) : range.GetOffset(x, y);
                        if((ignore == 1 || ignore == 3) && v == null)
                        {
                            continue;
                        }
                        else if((ignore == 2 || ignore == 3) && ExcelErrorValue.IsErrorValue(v?.ToString()))
                        {
                            continue;
                        }
                        result.Add(v);
                    }
                }
                var resultRange = new InMemoryRange(new RangeDefinition(1, (short)result.Count));
                var col = 0;
                foreach(var val in result)
                {
                    resultRange.SetValue(0, col++, val);
                }
                return CreateResult(resultRange, DataType.ExcelRange);
            }
            return CompileResultFactory.Create(firstArg.Value);
        }
    }
}
