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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
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
        private readonly ValueMatcher _valueMatcher = new WildCardValueMatcher();
        private enum MatchMode : int
        {
            ExactMatch = 0,
            ExactMatchReturnNextSmaller = -1,
            ExactMatchReturnNextLarger = 1,
            Wildcard = 2
        }

        private enum SearchMode : int
        {
            StartingAtFirst = 1,
            ReverseStartingAtLast = -1,
            BinarySearchAscending = 2,
            BinarySearchDescending = 3
        }

        private MatchMode GetMatchMode(int mm)
        {
            switch(mm)
            {
                case 0:
                    return MatchMode.ExactMatch;
                case -1:
                    return MatchMode.ExactMatchReturnNextSmaller;
                case 1:
                    return MatchMode.ExactMatchReturnNextLarger;
                case 2:
                    return MatchMode.Wildcard;
                default:
                    throw new ArgumentException("Invalid match mode: " + mm.ToString());
            }
        }

        private SearchMode GetSearchMode(int sm)
        {
            switch(sm)
            {
                case 1:
                    return SearchMode.StartingAtFirst;
                case -1:
                    return SearchMode.ReverseStartingAtLast;
                case 2:
                    return SearchMode.BinarySearchAscending;
                case 3:
                    return SearchMode.BinarySearchDescending;
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

        private static object GetReturnValue(List<object> returnArray, Dictionary<object, List<int>> origIndexes, object candidate, SearchMode searchMode)
        {
            if(searchMode == SearchMode.ReverseStartingAtLast)
            {
                return returnArray[origIndexes[candidate].Last()];
            }
            else
            {
                return returnArray[origIndexes[candidate].First()];
            }
        }

        private Dictionary<object, List<int>> CreateIndexes(List<object> lookupArray)
        {
            var origIndexes = new Dictionary<object, List<int>>();
            for (var i = 0; i < lookupArray.Count; i++)
            {
                if (!origIndexes.ContainsKey(lookupArray[i]))
                {
                    origIndexes.Add(lookupArray[i], new List<int>());
                }
                origIndexes[lookupArray[i]].Add(i);
            }
            return origIndexes;
        }

        private object GetSearchedValue(object lookupValue, List<object> lookupArray, List<object > returnArray, MatchMode matchMode, SearchMode searchMode)
        {
            if(matchMode == MatchMode.ExactMatch || lookupArray.IndexOf(lookupValue) > -1)
            {
                if (searchMode == SearchMode.ReverseStartingAtLast)
                {
                    return returnArray[lookupArray.LastIndexOf(lookupValue)];
                }
                return returnArray[lookupArray.IndexOf(lookupValue)];
            }
            var origIndexes = CreateIndexes(lookupArray);
            lookupArray.Sort((a, b) => {
                if (a == null && b != null) return 1.CompareTo(2);
                if (a != null && b == null) return 2.CompareTo(1);
                return CompareObjects(a, b);
            });
            if(matchMode == MatchMode.ExactMatchReturnNextSmaller)
            {
                var ix = searchMode == SearchMode.ReverseStartingAtLast ? returnArray.Count - 1 : 0;
                var prev = default(object);
                while(searchMode == SearchMode.ReverseStartingAtLast ? ix >= 0 : ix < returnArray.Count)
                {
                    var candidate = lookupArray[ix];
                    var res = CompareObjects(lookupValue, candidate);
                    if (res == 1)
                    {
                        prev = candidate;
                    }
                    else if(res == 0)
                    {
                        return GetReturnValue(returnArray, origIndexes, candidate, searchMode);
                    }
                    else
                    {
                        return GetReturnValue(returnArray, origIndexes, prev, searchMode);
                    }
                    if(searchMode == SearchMode.ReverseStartingAtLast)
                    {
                        ix--;
                    }
                    else
                    {
                        ix++;
                    }
                }
            }
            else if (matchMode == MatchMode.ExactMatchReturnNextLarger)
            {
                var ix = 0;
                while (ix < returnArray.Count)
                {
                    var candidate = lookupArray[ix];
                    var next = default(object);
                    if(ix < returnArray.Count - 2)
                    {
                        next = lookupArray[ix + 1];
                    }
                    var res = CompareObjects(lookupValue, candidate);
                    if (res == 0)
                    {
                        return GetReturnValue(returnArray, origIndexes, candidate, searchMode);
                    }
                    else if(next != null && res == 1)
                    {
                        var nextRes = CompareObjects(lookupValue, next);
                        if(nextRes == -1 || nextRes == 0)
                        {
                            return GetReturnValue(returnArray, origIndexes, next, searchMode);
                        }
                    }
                    ix++;
                }
                return null;
            }
            else if(matchMode == MatchMode.Wildcard)
            {
                var ix = searchMode == SearchMode.ReverseStartingAtLast ? returnArray.Count - 1 : 0;
                while (searchMode == SearchMode.ReverseStartingAtLast ? ix >= 0 : ix < returnArray.Count)
                {
                    var candidate = lookupArray[ix];
                    if(_valueMatcher.IsMatch(lookupValue, candidate) == 0)
                    {
                        return GetReturnValue(returnArray, origIndexes, candidate, searchMode);
                    }
                    if (searchMode == SearchMode.ReverseStartingAtLast)
                    {
                        ix--;
                    }
                    else
                    {
                        ix++;
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
            if(arguments.Count() > 3 && arguments.ElementAt(3) != null)
            {
                notFoundText = ArgToString(arguments, 3);
            }
            var matchMode = MatchMode.ExactMatch;
            if(arguments.Count() > 4 && arguments.ElementAt(4) != null)
            {
                var mm = ArgToInt(arguments, 4);
                matchMode = GetMatchMode(mm);
            }
            var searchMode = SearchMode.StartingAtFirst;
            if(arguments.Count() > 5)
            {
                var sm = ArgToInt(arguments, 5);
                searchMode = GetSearchMode(sm);
            }
            if(lookupArray.IndexOf(lookupValue) < 0 && matchMode == MatchMode.ExactMatch)
            {
                return string.IsNullOrEmpty(notFoundText) ? CreateResult(eErrorType.NA) : CreateResult(notFoundText, DataType.String);
            }
            var result = GetSearchedValue(lookupValue, lookupArray, returnArray, matchMode, searchMode);
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", sw.ElapsedMilliseconds);
            }
            return CreateResult(result, DataType.Unknown);
        }
    }
}
