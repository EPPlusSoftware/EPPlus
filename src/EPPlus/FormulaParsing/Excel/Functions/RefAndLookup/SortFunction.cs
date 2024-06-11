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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
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
        Description = "Sorts the contents of a range or array in ascending or descending order and returns a dynamic array.",
        SupportsArrays = true)]
    internal class SortFunction : ExcelFunction
    {        
        private readonly InMemoryRangeSorter _sorter = new InMemoryRangeSorter();
        public override string NamespacePrefix => "_xlfn._xlws.";
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            if(!arg1.IsExcelRange)
            {
                return CompileResultFactory.CreateDynamicArray(arg1.Value);
            }
            var range = arg1.ValueAsRangeInfo;
            var rangeDef = new RangeDefinition(range.Size.NumberOfRows, range.Size.NumberOfCols);
            var sortIndex = 1;
            if(arguments.Count > 1)
            {
                sortIndex = ArgToInt(arguments, 1, out ExcelErrorValue e1, 1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            }
            var sortOrder = 1;
            if(arguments.Count > 2)
            {
                sortOrder = ArgToInt(arguments, 2, out ExcelErrorValue e2, 1);
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            }
            var byCol = false;
            if(arguments.Count > 3)
            {
                byCol = ArgToBool(arguments, 3, false);
            }

            //Validate
            var maxIndex = byCol ? range.Size.NumberOfCols : range.Size.NumberOfRows;
            if (sortIndex < 1 || sortIndex > maxIndex) return CreateResult(eErrorType.Value);
            if (sortOrder != -1 && sortOrder != 1) return CreateResult(eErrorType.Value);
            var sortedRange = GetSortedRange(range, sortIndex, sortOrder, byCol);
            return CreateDynamicArrayResult(sortedRange, DataType.ExcelRange);
        }

        private InMemoryRange GetSortedRange(IRangeInfo sourceRange, int sortIndex, int sortOrder, bool byCol)
        {
            return byCol ?
                _sorter.SortByCol(sourceRange, sortIndex, sortOrder) :
                _sorter.SortByRow(sourceRange, sortIndex, sortOrder);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
