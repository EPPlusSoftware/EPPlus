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
        Description = "Returns the specified columns from an array.")]
    internal class ChooseCols : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments.First();
            var cols = new List<int>();
            for(var x = 1; x < arguments.Count(); x++)
            {
                var c = ArgToInt(arguments, x, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                cols.Add(c);
            }
            if(firstArg.IsExcelRange)
            {
                var source = firstArg.ValueAsRangeInfo;
                if (cols.Any(c => Math.Abs(c - 1) > source.Size.NumberOfCols || c == 0))
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                var nRows = source.Size.NumberOfRows;
                var resultRange = new InMemoryRange(new RangeDefinition(nRows, (short)cols.Count));
                var cIx = 0;
                foreach(var col in cols)
                {
                    for(var row = 0; row < nRows; row++)
                    {
                        var sourceIx = col > 0 ? col - 1 : source.Size.NumberOfCols + col;

                        var val = source.GetOffset(row, sourceIx);
                        resultRange.SetValue(row, cIx, val);
                    }
                    cIx++;
                }
                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
            else if(!cols.Any(x => x > 1))
            {
                var resultRange = new InMemoryRange(new RangeDefinition(1, (short)cols.Count));
                var cIx = 0;
                foreach(var col in cols)
                {
                    resultRange.SetValue(0, cIx++, firstArg.Value);
                }
                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
            return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
