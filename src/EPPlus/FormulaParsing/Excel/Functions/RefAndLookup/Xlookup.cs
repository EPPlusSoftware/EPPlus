/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.LookupAndReference,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2016",
            Description = "Searches a range or an array, and then returns the item corresponding to the first match it finds. Will return a VALUE error if the functions returns an array (EPPlus does not support dynamic arrayformulas)")]
    internal class Xlookup : LookupFunction
    {
        private enum SearchMode : int
        {
            ExactMatch = 0,
            ExactMatchReturnNextSmaller = -1,
            ExactMatchReturnNextLarger = 1,
            Wildcard = 2
        }

        private SearchMode GetSearchMode(int sm)
        {
            switch(sm)
            {
                case 0:
                    return SearchMode.ExactMatch;
                case -1:
                    return SearchMode.ExactMatchReturnNextSmaller;
                case 1:
                    return SearchMode.ExactMatchReturnNextLarger;
                case 2:
                    return SearchMode.Wildcard;
                default:
                    throw new ArgumentException("Invalid search mode: " + sm.ToString());
            }
        }

        protected int CompareObjects(object x1, object y1)
        {
            int ret;
            var isNumX = ConvertUtil.IsNumericOrDate(x1);
            var isNumY = ConvertUtil.IsNumericOrDate(y1);
            if (isNumX && isNumY)   //Numeric Compare
            {
                var d1 = ConvertUtil.GetValueDouble(x1);
                var d2 = ConvertUtil.GetValueDouble(y1);
                if (double.IsNaN(d1))
                {
                    d1 = double.MaxValue;
                }
                if (double.IsNaN(d2))
                {
                    d2 = double.MaxValue;
                }
                ret = d1 < d2 ? -1 : (d1 > d2 ? 1 : 0);
            }
            else if (isNumX == false && isNumY == false)   //String Compare
            {
                var s1 = x1 == null ? "" : x1.ToString();
                var s2 = y1 == null ? "" : y1.ToString();
                ret = string.Compare(s1, s2, StringComparison.CurrentCulture);
            }
            else
            {
                ret = isNumX ? -1 : 1;
            }

            return ret;
        }

        private object GetSearchedValue(object lookupValue, List<object> lookupArray, List<object > returnArray, SearchMode searchMode)
        {
            if(searchMode == SearchMode.ExactMatch || lookupArray.IndexOf(lookupValue) > -1)
            {
                return returnArray[lookupArray.IndexOf(lookupValue)];
            }
            var origIndexes = new Dictionary<object, int>();
            for(var i = 0; i < lookupArray.Count;i++)
            {
                if(!origIndexes.ContainsKey(lookupArray[i]))
                    origIndexes[lookupArray[i]] = i;
            }
            lookupArray.Sort((a, b) => {
                if (a == null && b != null) return 1.CompareTo(2);
                if (a != null && b == null) return 2.CompareTo(1);
                return CompareObjects(a, b);
            });
            if(searchMode == SearchMode.ExactMatchReturnNextSmaller)
            {
                var ix = 0;
                var prev = default(object);
                while(ix++ < returnArray.Count)
                {
                    var candidate = lookupArray[ix];
                    var res = CompareObjects(lookupValue, candidate);
                    if (res == -1)
                    {
                        prev = candidate;
                    }
                    else
                    {
                        return returnArray[origIndexes[candidate]];
                    }
                }
            }
            return null;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            Stopwatch sw = null;
            if (context.Debug)
            {
                sw = new Stopwatch();
                sw.Start();
            }
            ValidateArguments(arguments, 3);
            var lookupValue = arguments.ElementAt(0).Value;
            var lookupArray = Enumerable.Empty<object>().ToList();
            if(arguments.ElementAt(1).IsExcelRange)
            {
                lookupArray = arguments.ElementAt(1).ValueAsRangeInfo.Select(x => x.Value).ToList();
            }
            else
            {
                lookupArray = ArgsToObjectEnumerable(true, new List<FunctionArgument> { arguments.ElementAt(1) }, context).ToList();
            }
            var returnArray = Enumerable.Empty<object>().ToList();
            if (arguments.ElementAt(1).IsExcelRange)
            {
                returnArray = arguments.ElementAt(2).ValueAsRangeInfo.Select(x => x.Value).ToList();
            }
            else
            {
                returnArray = ArgsToObjectEnumerable(true, new List<FunctionArgument> { arguments.ElementAt(2) }, context).ToList();
            }
            var notFoundText = string.Empty;
            if(arguments.Count() > 3)
            {
                notFoundText = ArgToString(arguments, 3);
            }
            var searchMode = SearchMode.ExactMatch;
            if(arguments.Count() > 4)
            {
                var sm = ArgToInt(arguments, 4);
                searchMode = GetSearchMode(sm);
            }
            if(lookupArray.IndexOf(lookupValue) < 0 && searchMode == SearchMode.ExactMatch)
            {
                return string.IsNullOrEmpty(notFoundText) ? CreateResult(eErrorType.NA) : CreateResult(notFoundText, DataType.String);
            }
            var result = GetSearchedValue(lookupValue, lookupArray, returnArray, searchMode);
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", sw.ElapsedMilliseconds);
            }
            return CreateResult(result, DataType.Unknown);
        }
    }
}
