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
        Description = "Looks up a supplied value in the first row of a table, and returns the corresponding value from another row",
        SupportsArrays = true)]
    internal class HLookupV2 : ExcelFunction
    {
        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var searchedValue = arguments.ElementAt(0).Value;
            var arg1 = arguments.ElementAt(1);
            if (arg1.DataType == DataType.ExcelError) return CompileResult.GetErrorResult(((ExcelErrorValue)arg1.Value).Type);
            var lookupRange = arg1.ValueAsRangeInfo;
            var lookupIndex = ArgToInt(arguments, 2);
            var rangeLookup = true;
            if (arguments.Count() > 3)
            {
                rangeLookup = ArgToBool(arguments, 3);
            }
            var index = -1;
            if (!rangeLookup)
            {
                var scanner = new XlookupScanner(searchedValue, lookupRange, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchWithWildcard, LookupRangeDirection.Horizontal);
                index = scanner.FindIndex();
                if (index < 0)
                {
                    return CreateResult(eErrorType.NA);
                }
            }
            else
            {
                index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, true, new LookupComparer(LookupMatchMode.ExactMatchReturnNextSmaller), LookupRangeDirection.Horizontal);
                index = LookupBinarySearch.GetMatchIndex(index, lookupRange, LookupMatchMode.ExactMatchReturnNextSmaller, true);
                if (index < 0)
                {
                    return CreateResult(eErrorType.NA);
                }
            }
            return CompileResultFactory.Create(lookupRange.GetOffset(lookupIndex - 1, index));
        }
    }
}
