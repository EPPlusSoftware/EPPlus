using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Searches for a specific value in one data vector, and returns a value from the corresponding position of a second data vector",
        SupportsArrays = true)]
    internal class Lookup : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments[0].Value ?? 0;     //If Search value is null, we should search for 0 instead
            var arg2 = arguments[1];
            if(!arg2.IsExcelRange)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var lookupRange = arg2.ValueAsRangeInfo;
            var returnVector = lookupRange;
            var separateReturnVector = false;
            if(arguments.Count > 2 && arguments[2].IsExcelRange)
            {
                separateReturnVector = true;
                returnVector = arguments[2].ValueAsRangeInfo;
            }
            var nLookupRows = lookupRange.Size.NumberOfRows;
            var nLookupCols = lookupRange.Size.NumberOfCols;
            var index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, true, new LookupComparer(LookupMatchMode.ExactMatchReturnNextSmaller));
            index = LookupBinarySearch.GetMatchIndex(index, returnVector, LookupMatchMode.ExactMatchReturnNextSmaller, true);
            if(index < 0)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            var nReturnRows = returnVector.Size.NumberOfRows;
            var nReturnCols = returnVector.Size.NumberOfCols;
            if(nReturnRows >= nReturnCols)
            {
                if(separateReturnVector && nReturnCols > 1)
                {
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
                return CompileResultFactory.Create(returnVector.GetOffset(index, nReturnCols - 1));
            }
            else
            {
                if (separateReturnVector && nReturnRows > 1)
                {
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
                return CompileResultFactory.Create(returnVector.GetOffset(nReturnRows - 1, index));
            }
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
