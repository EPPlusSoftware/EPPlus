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
        Description = "Looks up a supplied value in the first column of a table, and returns the corresponding value from another column",
        SupportsArrays = true)]
    internal class VLookupV2 : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 3;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 1)
            {
                return FunctionParameterInformation.IgnoreAddress;
            }
            return FunctionParameterInformation.Normal;
        }));
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments[0].Value ?? 0;     //If Search value is null, we should search for 0 instead
            var arg1 = arguments[1];
            if (arg1.DataType == DataType.ExcelError) return CompileResult.GetErrorResult(((ExcelErrorValue)arg1.Value).Type);
            var lookupRange = arg1.GetAsRangeInfo(context);
            var lookupIndex = ArgToInt(arguments, 2);
            var rangeLookup = true;
            if(arguments.Count() > 3)
            {
                rangeLookup = ArgToBool(arguments, 3);
            }
            int index;
            if (!rangeLookup)
            {
                var scanner = new XlookupScanner(searchedValue, lookupRange, LookupSearchMode.StartingAtFirst, LookupMatchMode.ExactMatchWithWildcard, LookupRangeDirection.Vertical);
                index = scanner.FindIndex();
                if (index < 0)
                {
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
            }
            else
            {
                index = LookupBinarySearch.BinarySearch(searchedValue, lookupRange, true, new LookupComparer(LookupMatchMode.ExactMatchReturnNextSmaller), LookupRangeDirection.Vertical);
                index = LookupBinarySearch.GetMatchIndex(index, lookupRange, LookupMatchMode.ExactMatchReturnNextSmaller, true);
                if (index < 0)
                {
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
            }
            return CompileResultFactory.Create(lookupRange.GetOffset(index, lookupIndex - 1));
        }
    }
}
