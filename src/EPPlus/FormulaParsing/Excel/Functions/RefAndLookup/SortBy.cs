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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Sorts the contents of a range or array based on the values in a corresponding range or array.",
        SupportsArrays = true)]
    internal class SortBy : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = ArgToRangeInfo(arguments, 0);
            var nArgs = arguments.Count();
            var nRows = range.Size.NumberOfRows;
            var nCols = range.Size.NumberOfCols;
            var byRanges = new List<IRangeInfo>();
            var sortOrders = new List<short>();
            var direction = LookupDirection.Vertical;
            for(var x = 1; x < nArgs; x+=2)
            {
                var byRange = ArgToRangeInfo(arguments, x);
                if(x == 1)
                {
                    if(byRange.Size.NumberOfCols > byRange.Size.NumberOfRows)
                    {
                        direction = LookupDirection.Horizontal;
                    }
                }
                else if(
                    (direction == LookupDirection.Horizontal && byRange.Size.NumberOfRows > byRange.Size.NumberOfCols) 
                    ||
                    (direction == LookupDirection.Vertical && byRange.Size.NumberOfCols > byRange.Size.NumberOfRows))
                {
                    // two "by-ranges" goes in different direction (i.e. vertical/horizontal) which is not allowed.
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                if (byRange.Size.NumberOfCols != nCols && byRange.Size.NumberOfRows != nRows)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                if(byRange.Size.NumberOfRows > 1 && byRange.Size.NumberOfCols > 1)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                byRanges.Add(byRange);
                var sortOrder = 1;
                if(x + 1 < nArgs)
                {
                    sortOrder = ArgToInt(arguments, x + 1, out ExcelErrorValue e1, 1);
                    if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                    if(sortOrder != 1 && sortOrder != -1)
                    {
                        return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                    }
                }
                sortOrders.Add((short)sortOrder);
            }
            var sortByImpl = new SortByImpl(range, byRanges, sortOrders, direction);
            var sortedRange = sortByImpl.Sort();
            return CreateDynamicArrayResult(sortedRange, DataType.ExcelRange);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
