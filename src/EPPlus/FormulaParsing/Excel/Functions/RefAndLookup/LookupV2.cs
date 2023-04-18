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
    internal class LookupV2 : ExcelFunction
    {
        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var searchedValue = arguments.ElementAt(0).Value ?? 0;     //If Search value is null, we should search for 0 instead
            var arg2 = arguments.ElementAt(1);
            if(!arg2.IsExcelRange)
            {
                return CreateResult(eErrorType.Value);
            }
            var lookupRange = arg2.ValueAsRangeInfo;
            var returnVector = lookupRange;
            var separateReturnVector = false;
            if(arguments.Count() > 2 && arguments.ElementAt(2).IsExcelRange)
            {
                separateReturnVector = true;
                returnVector = arguments.ElementAt(2).ValueAsRangeInfo;
            }
            var nLookupRows = lookupRange.Size.NumberOfRows;
            var nLookupCols = lookupRange.Size.NumberOfCols;
            var index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, true, new LookupComparer(LookupMatchMode.ExactMatchReturnNextSmaller));
            index = LookupBinarySearch.GetMatchIndex(index, returnVector, LookupMatchMode.ExactMatchReturnNextSmaller, true);
            if(index < 0)
            {
                return CreateResult(eErrorType.NA);
            }
            var nReturnRows = returnVector.Size.NumberOfRows;
            var nReturnCols = returnVector.Size.NumberOfCols;
            if(nReturnRows >= nReturnCols)
            {
                if(separateReturnVector && nReturnCols > 1)
                {
                    return CreateResult(eErrorType.NA);
                }
                return CompileResultFactory.Create(returnVector.GetOffset(index, nReturnCols - 1));
            }
            else
            {
                if (separateReturnVector && nReturnRows > 1)
                {
                    return CreateResult(eErrorType.NA);
                }
                return CompileResultFactory.Create(returnVector.GetOffset(nReturnRows - 1, index));
            }
        }
    }
}
