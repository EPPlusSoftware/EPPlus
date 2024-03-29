﻿/*************************************************************************************************
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
        IntroducedInExcelVersion = "2021",
        SupportsArrays = true)]
    internal class Sequence : ExcelFunction
    {

        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rows = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var argCount = arguments.Count;
            var columns = 1;
            if(argCount > 1)
            {
                columns = ArgToInt(arguments, 1, out ExcelErrorValue e2);
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            }
            var start = 1d;
            if(argCount > 2)
            {
                start = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
                if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            }
            var step = 1d;
            if(argCount > 3)
            {
                step = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
                if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            } 
            
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
