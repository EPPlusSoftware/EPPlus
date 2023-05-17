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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "7",
        Description = "Returns an array with a sequence of numbers",
        IntroducedInExcelVersion = "2021")]
    internal class Sequence : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var rows = ArgToInt(arguments, 0); 
            var argCount = arguments.Count();
            var columns = argCount > 1 ? ArgToInt(arguments, 1) : 1;
            var start = argCount > 2 ? ArgToDecimal(arguments, 2) : 1;
            var step = argCount > 3 ? ArgToDecimal(arguments, 3) : 1;
            
            if (rows<0 || columns < 0)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            else if(rows==0 || columns==0)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
            }

            var size = new RangeDefinition(rows, (short)columns);
            var range = new InMemoryRange(size);

            SetSequence(range, start, step);

            return CreateDynamicArrayResult(range, DataType.ExcelRange);
        }

        private void SetSequence(InMemoryRange range, double start, double step)
        {
            var v = start;
            for (int r = 0; r < range.Size.NumberOfRows; r++)
            {
                for (int c = 0; c < range.Size.NumberOfCols; c++)
                {
                    range.SetValue(r, c, v);
                    v += step;
                }
            }
        }
    }
}
