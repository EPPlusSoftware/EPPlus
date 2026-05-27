/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  19/3/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions
{
    internal abstract class GroupByFunctionBase : ExcelFunction
    {
        protected readonly LookupComparerBase _comparer = new SortByComparer();

        protected const int TotalDepthNoTotals = 0;
        protected const int TotalDepthGrandOnly = 1;

        private InMemoryRange _allValuesRangeCache;
        private List<object[]> _allValuesCacheKey;

        protected List<string> ResolveFunctionHeaders(List<LambdaCalculator> functions)
        {
            var names = functions
                .Select(f => f.EtaFunction != null ? f.EtaFunction.Name : "CUSTOM")
                .ToList();

            int customCount = names.Count(n => n == "CUSTOM");

            if (customCount > 1)
            {
                int counter = 1;
                for (int i = 0; i < names.Count; i++)
                    if (names[i] == "CUSTOM")
                        names[i] = $"CUSTOM{counter++}";
            }

            return names;
        }

        // -------------------------------------------------------
        // Argument parsing 
        // -------------------------------------------------------

        protected bool Fail(eErrorType err, out CompileResult error)
        {
            error = CompileResult.GetErrorResult(err);
            return false;
        }

        protected bool TryParseFunctionArg(FunctionArgument funtionArgument, List<LambdaCalculator> functions,
                                            out LambdaCalculator function, out FunctionLayout layout)
        {
            function = null;
            layout = FunctionLayout.Single;
            if (funtionArgument.DataType == DataType.LambdaCalculation)
            {
                // Single function
                function = funtionArgument.Value as LambdaCalculator;
                functions.Add(function);
            }
            else if (funtionArgument.IsExcelRange)
            {
                // Multiple functions via HSTACK or VSTACK
                var range = funtionArgument.ValueAsRangeInfo;
                bool isHorizontal = range.Size.NumberOfRows == 1;
                layout = isHorizontal ? FunctionLayout.Horizontal : FunctionLayout.Vertical;

                int count = isHorizontal ? range.Size.NumberOfCols : range.Size.NumberOfRows;
                for (int i = 0; i < count; i++)
                {
                    var cellVal = isHorizontal
                        ? range.GetOffset(0, i)
                        : range.GetOffset(i, 0);

                    if (cellVal is LambdaCalculator lc)
                        functions.Add(lc);
                    else
                        return false;
                }
                function = functions[0];
            }
            else
            {
                return false;
            }
            return true;
        }

        protected bool TryParseTotalDepthArg(FunctionArgument arg, int numberOfCols,
            out int totalDepth)
        {
            totalDepth = Convert.ToInt32(arg.Value);
            if (Math.Abs(totalDepth) > numberOfCols)
                return false;
            return true;
        }

        protected int[] ParseSortOrderArg(FunctionArgument arg)
        {
            if (arg.IsExcelRange)
            {
                var range = arg.ValueAsRangeInfo;
                bool isHorizontal = range.Size.NumberOfRows == 1;
                int count = isHorizontal ? range.Size.NumberOfCols : range.Size.NumberOfRows;
                var result = new int[count];
                for (int i = 0; i < count; i++)
                    result[i] = Convert.ToInt32(isHorizontal
                        ? range.GetOffset(0, i)
                        : range.GetOffset(i, 0));
                return result;
            }
            else
            {
                return new[] { Convert.ToInt32(arg.Value) };
            }
        }

        // -------------------------------------------------------
        // Header resolution
        // -------------------------------------------------------
        protected FieldHeaders ResolveHeaders(FieldHeaders headers, IRangeInfo values)
        {
            if (headers != FieldHeaders.Missing)
                return headers;

            if (values.Size.NumberOfRows < 2)
                return FieldHeaders.No;

            var first = values.GetOffset(0, 0);
            var second = values.GetOffset(1, 0);

            bool firstIsText = first is string;
            bool secondIsNumber = second is double || second is int || second is long || second is float;

            return firstIsText && secondIsNumber
                ? FieldHeaders.YesAndDontShow
                : FieldHeaders.No;
        }

        // -------------------------------------------------------
        // Grouping
        // -------------------------------------------------------
        protected List<GroupLevel> BuildGroups(GroupByBaseArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            bool multipleFunctions = args.Functions.Count > 1;
            int startRow = hasHeaders ? 1 : 0;

            int nKeyCols = args.RowFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;

            var rootDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
            var rootOrder = new List<string>();

            for (int r = startRow; r < args.RowFields.Size.NumberOfRows; r++)
            {
                if (args.FilterArray != null)
                {
                    var filterVal = args.FilterArray.GetOffset(r, 0);
                    if (filterVal is bool b && !b) continue;
                    if (filterVal is int i && i == 0) continue;
                }

                var keyParts = new object[nKeyCols];
                for (int c = 0; c < nKeyCols; c++)
                    keyParts[c] = args.RowFields.GetOffset(r, c);

                var rowVals = new object[nValCols];
                for (int c = 0; c < nValCols; c++)
                    rowVals[c] = args.Values.GetOffset(r, c);

                var currentDict = rootDict;
                var currentOrder = rootOrder;
                GroupLevel currentLevel = null;

                for (int depth = 0; depth < nKeyCols; depth++)
                {
                    var keyStr = (keyParts[depth]?.ToString() ?? string.Empty).ToLowerInvariant();
                    if (!currentDict.TryGetValue(keyStr, out currentLevel))
                    {
                        currentLevel = new GroupLevel { Key = keyParts[depth] };
                        currentDict[keyStr] = currentLevel;
                        currentOrder.Add(keyStr);
                    }

                    if (depth < nKeyCols - 1) // If there are more cols after the current, add child.
                    {
                        if (currentLevel.ChildDict == null)
                        {
                            currentLevel.ChildDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
                            currentLevel.ChildOrder = new List<string>();
                        }
                        currentDict = currentLevel.ChildDict;
                        currentOrder = currentLevel.ChildOrder;
                    }
                }

                var leafKey = string.Join("|", keyParts.Select(k => k?.ToString().ToLowerInvariant() ?? string.Empty).ToArray());
                var row = currentLevel.Rows.FirstOrDefault(rw =>
                    string.Join("|", rw.KeyParts.Select(k => k?.ToString().ToLowerInvariant() ?? string.Empty).ToArray()) == leafKey);

                if (row == null)
                {
                    row = new GroupRow { KeyParts = keyParts };
                    currentLevel.Rows.Add(row);
                }
                row.Values.Add(rowVals);
                args.AllValuesInOrder.Add(rowVals);
            }

            var levels = BuildOrderedTree(rootDict, rootOrder);
            AggregateTree(levels, args, context);
            return levels;
        }



        protected List<GroupLevel> BuildOrderedTree(
            Dictionary<string, GroupLevel> dict,
            List<string> order)
        {
            var levels = order.Select(k => dict[k]).ToList();
            foreach (var level in levels)
                if (level.ChildDict != null)
                    level.Children = BuildOrderedTree(level.ChildDict, level.ChildOrder);
            return levels;
        }

        protected void AggregateTree(List<GroupLevel> levels, GroupByBaseArgs args, ParsingContext context)
        {
            foreach (var level in levels)
            {
                if (level.IsLeaf)
                {
                    foreach (var row in level.Rows)
                    {
                        row.AggregatedValues = args.Functions.Select(f =>
                        {
                            int nValCols = row.Values[0].Length;
                            var result = new object[nValCols];
                            for (int col = 0; col < nValCols; col++)
                            {
                                result[col] = Aggregate(f, row.Values, col, context,
                                    f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                            }
                            return result;
                        }).ToList();
                        row.AggregatedValue = row.AggregatedValues[0][0];
                    }

                    var allVals = level.Rows.SelectMany(r => r.Values).ToList();
                    level.SubtotalValues = args.Functions.Select(f =>
                    {
                        int nValCols = allVals[0].Length;
                        var result = new object[nValCols];
                        for (int col = 0; col < nValCols; col++)
                        {
                            result[col] = Aggregate(f, allVals, col, context,
                                f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                        }
                        return result;
                    }).ToList();
                    level.SubtotalValue = level.SubtotalValues[0][0];
                }
                else
                {
                    AggregateTree(level.Children, args, context);

                    var allVals = level.Children.SelectMany(c => GetAllValues(c)).ToList();
                    level.SubtotalValues = args.Functions.Select(f =>
                    {
                        int nValCols = allVals[0].Length;
                        var result = new object[nValCols];
                        for (int col = 0; col < nValCols; col++)
                        {
                            result[col] = Aggregate(f, allVals, col, context,
                                f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                        }
                        return result;
                    }).ToList();
                    level.SubtotalValue = level.SubtotalValues[0][0];
                }
            }
        }

        protected List<object[]> GetAllValues(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows.SelectMany(r => r.Values).ToList();
            return level.Children.SelectMany(c => GetAllValues(c)).ToList();
        }

        private static InMemoryRange BuildRangeFromList(List<object[]> values)
        {
            int nRows = values.Count;
            int nCols = values.Count > 0 ? values[0].Length : 1;
            var range = new InMemoryRange(nRows, (short)nCols);
            for (int row = 0; row < nRows; row++)
                for (int col = 0; col < nCols; col++)
                    range.SetValue(row, col, values[row][col]);
            return range;
        }

        protected object Aggregate(LambdaCalculator calculator, List<object[]> values, ParsingContext context, List<object[]> allValues = null)
        {
            var range = BuildRangeFromList(values);

            calculator.BeginCalculation();
            calculator.SetVariableValue(0, range, DataType.ExcelRange, context);

            if (calculator.NumberOfVariables > 1 && allValues != null)
            {
                // Cacha: om vi får samma allValues-referens som förra gången,
                // återanvänd den InMemoryRange vi redan byggt.
                if (!ReferenceEquals(_allValuesCacheKey, allValues))
                {
                    _allValuesRangeCache = BuildRangeFromList(allValues);
                    _allValuesCacheKey = allValues;
                }
                calculator.SetVariableValue(1, _allValuesRangeCache, DataType.ExcelRange, context);
            }
            return calculator.Execute(context).ResultValue;
        }

        protected object Aggregate(LambdaCalculator calculator, List<object[]> values, int colIndex, ParsingContext context, List<object[]> allValues = null)
        {
            var range = BuildRangeFromColumn(values, colIndex);

            calculator.BeginCalculation();
            calculator.SetVariableValue(0, range, DataType.ExcelRange, context);

            if (calculator.NumberOfVariables > 1 && allValues != null)
            {
                if (!ReferenceEquals(_allValuesCacheKey, allValues))
                {
                    _allValuesRangeCache = BuildRangeFromList(allValues);
                    _allValuesCacheKey = allValues;
                }
                calculator.SetVariableValue(1, _allValuesRangeCache, DataType.ExcelRange, context);
            }
            return calculator.Execute(context).ResultValue;
        }

        protected List<GroupLevel> ApplySort(List<GroupLevel> levels, GroupByBaseArgs args, int depth = 1)
        {
            if (args.SortOrders == null || args.SortOrders.All(s => s == 0)) return levels;

            if (args.FieldRelationship == FieldRelationship.Table)
            {
                var allRows = levels.SelectMany(l => CollectLeafRows(l)).ToList();
                allRows = SortRowsMulti(allRows, args);

                var newLevelDict = new Dictionary<string, GroupLevel>();
                var newLevelOrder = new List<string>();
                foreach (var row in allRows)
                {
                    var topKey = (row.KeyParts[0]?.ToString() ?? string.Empty).ToLowerInvariant();
                    if (!newLevelDict.TryGetValue(topKey, out var level))
                    {
                        level = new GroupLevel { Key = row.KeyParts[0] };
                        newLevelDict[topKey] = level;
                        newLevelOrder.Add(topKey);
                    }
                    level.Rows.Add(row);
                }
                return newLevelOrder.Select(k => newLevelDict[k]).ToList();
            }
            else
            {
                var sortForThisLevel = args.SortOrders
                    .FirstOrDefault(s => Math.Abs(s) == depth);

                bool hasSortForThisLevel = sortForThisLevel != 0;
                bool desc = sortForThisLevel < 0;
                bool sortOnAggregated = hasSortForThisLevel && Math.Abs(sortForThisLevel) > args.RowFields.Size.NumberOfCols;

                if (hasSortForThisLevel)
                {
                    levels = sortOnAggregated
                        ? (desc ? levels.OrderByDescending(l => l.SubtotalValue as IComparable, _comparer).ToList()
                                : levels.OrderBy(l => l.SubtotalValue as IComparable, _comparer).ToList())
                        : (desc ? levels.OrderByDescending(l => l.Key as IComparable, _comparer).ToList()
                                : levels.OrderBy(l => l.Key as IComparable, _comparer).ToList());
                }

                foreach (var level in levels)
                {
                    if (!level.IsLeaf)
                        level.Children = ApplySort(level.Children, args, depth + 1);
                    else
                        level.Rows = SortRowsMulti(level.Rows, args);
                }

                return levels;
            }
        }

        private List<GroupRow> SortRowsMulti(List<GroupRow> rows, GroupByBaseArgs args)
        {
            if (rows == null || rows.Count == 0) return rows;

            int nKeyCols = args.RowFields.Size.NumberOfCols;
            IOrderedEnumerable<GroupRow> ordered = null;

            foreach (var sortOrder in args.SortOrders)
            {
                if (sortOrder == 0) continue;
                bool desc = sortOrder < 0;
                int col = Math.Abs(sortOrder);
                bool sortOnAggregated = col > nKeyCols;

                // Capture loop variables
                var capturedCol = col;
                var capturedSortOnAggregated = sortOnAggregated;

                Func<GroupRow, object> keySelector = capturedSortOnAggregated
                    ? (Func<GroupRow, object>)(r => r.AggregatedValue)
                    : (r => r.KeyParts[Math.Min(capturedCol - 1, r.KeyParts.Length - 1)]);

                if (ordered == null)
                    ordered = desc
                        ? rows.OrderByDescending(keySelector, _comparer)
                        : rows.OrderBy(keySelector, _comparer);
                else
                    ordered = desc
                        ? ordered.ThenByDescending(keySelector, _comparer)
                        : ordered.ThenBy(keySelector, _comparer);
            }

            return ordered?.ToList() ?? rows;
        }

        private static InMemoryRange BuildRangeFromColumn(List<object[]> values, int colIndex)
        {
            int nRows = values.Count;
            var range = new InMemoryRange(nRows, 1);
            for (int row = 0; row < nRows; row++)
                range.SetValue(row, 0, values[row][colIndex]);
            return range;
        }

        private IEnumerable<GroupRow> CollectLeafRows(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows;
            return level.Children.SelectMany(c => CollectLeafRows(c));
        }
    }
}