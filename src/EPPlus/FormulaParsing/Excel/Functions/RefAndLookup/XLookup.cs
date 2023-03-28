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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RoundingHelper;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.LookupAndReference,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2016",
            Description = "Searches a range or an array, and then returns the item corresponding to the first match it finds.", 
            SupportsArrays = true)]
    internal class Xlookup : LookupFunction
    {

        private readonly ValueMatcher _valueMatcher = new WildCardValueMatcher();
        private Stopwatch _stopwatch = null;

        private int GetMatchIndex(object lookupValue, IRangeInfo lookupRange, IRangeInfo returnArray, bool asc, LookupMatchMode matchMode)
        {
            var comparer = new LookupComparer(matchMode);
            var ix = LookupBinarySearch.BinarySearch(lookupValue, lookupRange, asc, comparer);
            return LookupBinarySearch.GetMatchIndex(ix, returnArray, matchMode, asc);
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (context.Debug)
            {
                _stopwatch = new Stopwatch();
                _stopwatch.Start();
            }
            ValidateArguments(arguments, 3);
            var lookupValue = arguments.ElementAt(0).Value;

            // lookup range
            if (!arguments.ElementAt(1).IsExcelRange) return CreateResult(eErrorType.Value);
            var lookupRange = arguments.ElementAt(1).ValueAsRangeInfo;
            var lookupDirection = XlookupUtil.GetLookupDirection(lookupRange);
            if (lookupRange.Size.NumberOfRows > 1 && lookupRange.Size.NumberOfCols > 1) return CreateResult(eErrorType.Value);

            // return range
            if (!arguments.ElementAt(2).IsExcelRange) return CreateResult(eErrorType.Value);
            var returnArray = arguments.ElementAt(2).ValueAsRangeInfo;
            var notFoundText = string.Empty;

            // not found text
            if (arguments.Count() > 3 && arguments.ElementAt(3) != null)
            {
                notFoundText = ArgToString(arguments, 3);
            }

            // match mode
            var matchMode = LookupMatchMode.ExactMatch;
            if (arguments.Count() > 4 && arguments.ElementAt(4) != null)
            {
                var mm = ArgToInt(arguments, 4);
                matchMode = XlookupUtil.GetMatchMode(mm);
            }

            // search mode
            var searchMode = LookupSearchMode.StartingAtFirst;
            if (arguments.Count() > 5)
            {
                var sm = ArgToInt(arguments, 5);
                searchMode = XlookupUtil.GetSearchMode(sm);
            }
            int ix;
            if (searchMode == LookupSearchMode.BinarySearchAscending || searchMode == LookupSearchMode.BinarySearchDescending)
            {
                var asc = (searchMode == LookupSearchMode.BinarySearchAscending);
                ix = GetMatchIndex(lookupValue, lookupRange, returnArray, asc, matchMode);
            }
            else
            {
                var scanner = new XlookupScanner(lookupValue, lookupRange, searchMode, matchMode);
                ix = scanner.FindIndex();
            }
            if (context.Debug)
            {
                _stopwatch.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", _stopwatch.ElapsedMilliseconds);
            }
            return BuildCompileResult(lookupDirection, returnArray, notFoundText, ix);
        }

        private CompileResult BuildCompileResult(LookupRangeDirection lookupDirection, IRangeInfo returnArray, string notFoundText, int ix)
        {
            if (ix < 0 || ix > ((lookupDirection == LookupRangeDirection.Vertical) ? returnArray.Size.NumberOfRows - 1 : returnArray.Size.NumberOfCols - 1))
            {
                return string.IsNullOrEmpty(notFoundText) ? CreateResult(eErrorType.NA) : CreateResult(notFoundText, DataType.String);
            }
            var result = default(IRangeInfo);
            if (lookupDirection == LookupRangeDirection.Vertical)
            {
                var nCols = returnArray.Size.NumberOfCols;
                result = returnArray.GetOffset(ix, 0, ix, nCols - 1);
            }
            else
            {
                var nRows = returnArray.Size.NumberOfRows;
                result = returnArray.GetOffset(0, ix, nRows - 1, ix);
            }
            if (result == null)
            {
                if (string.IsNullOrEmpty(notFoundText))
                {
                    return CreateResult(eErrorType.NA);
                }
                else
                {
                    return CreateResult(notFoundText, DataType.String);
                }
            }
            return CreateResult(result, DataType.ExcelRange);
        }
    }
}
