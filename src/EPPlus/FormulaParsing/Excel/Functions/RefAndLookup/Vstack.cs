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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Combines arrays vertically into a single array.",
        SupportsArrays = true)]
    internal class Vstack : StackFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var ranges = GetRanges(arguments, out ExcelErrorValue err);
            if (err != null)
            {
                return CreateDynamicArrayResult(err, DataType.ExcelError);
            }
            var nRows = ranges.Sum(x => x.Size.NumberOfRows);
            var nCols = ranges.Max(x => x.Size.NumberOfCols);
            var range = new InMemoryRange(nRows, nCols);
            var rowIx = 0;
            foreach(var argRange in ranges)
            {
                for(var row = 0; row < argRange.Size.NumberOfRows;row++)
                {
                    var col = 0;
                    for(; col < argRange.Size.NumberOfCols;col++)
                    {
                        var v = argRange.GetOffset(row, col);
                        range.SetValue(rowIx, col, v);
                    }
                    for(;col < range.Size.NumberOfCols;col++)
                    {
                        range.SetValue(rowIx, col, ErrorValues.NAError);
                    }
                    rowIx++;
                }
            }
            return CreateDynamicArrayResult(range, DataType.ExcelRange);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
