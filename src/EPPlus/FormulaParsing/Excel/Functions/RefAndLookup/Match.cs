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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Finds the relative position of a value in a supplied array",
        SupportsArrays = true)]
    internal class Match : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments[0].Value ?? 0;     //If Search value is null, we should search for 0 instead
            var arg2 = arguments[1];
            if (!arg2.IsExcelRangeOrSingleCell) return CreateResult(eErrorType.NA);
            var lookupRange = arg2.ValueAsRangeInfo;
            var matchType = 1;
            if(arguments.Count > 2)
            {
                matchType = ArgToInt(arguments, 2, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                if (matchType > 1 || matchType < -1) return CreateResult(eErrorType.Value);
            }
            int index;
            if(matchType == 1)
            {
                index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, true, new LookupComparer(LookupMatchMode.ExactMatchReturnNextSmaller));
                index = LookupBinarySearch.GetMatchIndex(index, lookupRange, LookupMatchMode.ExactMatchReturnNextSmaller, true);
            }
            else if(matchType == -1)
            {
                index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, false, new LookupComparer(LookupMatchMode.ExactMatchReturnNextLarger));
                index = LookupBinarySearch.GetMatchIndex(index, lookupRange, LookupMatchMode.ExactMatchReturnNextLarger, false);
            }
            else
            {
                // matchType == 0...
                var scanner = new XlookupScanner(searchedValue, lookupRange, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchWithWildcard);
                index = scanner.FindIndex();
            }
            if(index < 0)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            return CreateResult(index + 1, DataType.Integer);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
