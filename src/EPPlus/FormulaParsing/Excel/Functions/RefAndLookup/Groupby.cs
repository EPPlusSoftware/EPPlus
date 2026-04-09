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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "",
        Description = "Allows you to create a summary of your data via a formula. Supports grouping along one axis and aggregating the associated values.")]

    internal class Groupby : GroupbyFunctionBase
    {
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParseBaseArgs(arguments, out var args, out var error))
                return error;
            var groups = BuildGroups(args, context);
            groups = ApplySort(groups, args);
            var result = BuildResult(groups, args, context);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        // -------------------------------------------------------
        // Sorting
        // -------------------------------------------------------
        private List<GroupLevel> ApplySort(List<GroupLevel> levels, GroupByBaseArgs args, int depth = 1)
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

        private IEnumerable<GroupRow> CollectLeafRows(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows;
            return level.Children.SelectMany(c => CollectLeafRows(c));
        }

        // -------------------------------------------------------
        // Build result
        // -------------------------------------------------------        

        private InMemoryRange BuildResult(List<GroupLevel> levels, GroupByBaseArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args);
            bool showHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                              || resolvedHeaders == FieldHeaders.NoButGenerate;
            bool addFunctionHeaders = args.Functions.Count > 1;
            bool showTotals = args.TotalDepth != TotalDepthNoTotals;
            bool totalsAtTop = args.TotalDepth < 0;
            bool totalsAtEnd = args.TotalDepth > 0;
            int subtotalDepth = Math.Abs(args.TotalDepth);
            bool showSubtotals = subtotalDepth > 1;
            bool grandAndSub = subtotalDepth > 1;

            int nKeyCols = args.RowFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;
            int nFunctions = args.Functions.Count;

            int valColsPerRow = args.FunctionLayout == FunctionLayout.Horizontal
                                ? nValCols * nFunctions
                                : args.FunctionLayout == FunctionLayout.Vertical
                                ? nValCols + 1  
                                : nValCols;
            int nCols = nKeyCols + valColsPerRow;

            int dataRows = CountDataRows(levels);
            int subtotalRows = showSubtotals
                ? CountSubtotalRows(levels, subtotalDepth, 1) * (args.FunctionLayout == FunctionLayout.Vertical ? nFunctions : 1)
                : 0;
            int resultDataRows = args.FunctionLayout == FunctionLayout.Vertical
                ? dataRows * nFunctions
                : dataRows;

            int grandTotalRows = showTotals
                ? (args.FunctionLayout == FunctionLayout.Vertical ? nFunctions : 1)
                : 0;

            int totalRows = resultDataRows + subtotalRows
                          + (showHeaders ? 1 : 0)
                          + grandTotalRows
                          + (addFunctionHeaders && args.FunctionLayout == FunctionLayout.Horizontal ? 1 : 0);

            var result = new InMemoryRange(totalRows, (short)nCols); 
            int r = 0;

            if(addFunctionHeaders)
            {
                var functionHeaders = ResolveFunctionHeaders(args);
                if(args.FunctionLayout == FunctionLayout.Horizontal)
                {
                    for (int c = 0; c < nFunctions; c++)
                        result.SetValue(r, c + 1, functionHeaders[c]);
                    r++;
                }                
            }
            if (showHeaders)
            {                
                for (int c = 0; c < nKeyCols; c++)
                    result.SetValue(r, c, resolvedHeaders == FieldHeaders.NoButGenerate
                        ? $"Field {c + 1}"
                        : args.RowFields.GetOffset(0, c)?.ToString());

                if (args.FunctionLayout == FunctionLayout.Horizontal)
                {
                    for (int f = 0; f < nFunctions; f++)
                        for (int c = 0; c < nValCols; c++)
                            result.SetValue(r, nKeyCols + f * nValCols + c, resolvedHeaders == FieldHeaders.NoButGenerate
                                ? $"Field {nKeyCols + f * nValCols + c + 1}"
                                : args.Values.GetOffset(0, c)?.ToString());
                }
                else
                {
                    for (int c = 0; c < nValCols; c++)
                        result.SetValue(r, nKeyCols + c + (addFunctionHeaders ? 1: 0), resolvedHeaders == FieldHeaders.NoButGenerate
                            ? $"Field {nKeyCols + c + 1}"
                            : args.Values.GetOffset(0, c)?.ToString());
                }
                r++;
            }

            string grandTotalStr = grandAndSub ? "Grand Total" : "Total";

            if (totalsAtTop && showTotals)
                r = WriteGrandTotal(result, r, levels, grandTotalStr, nKeyCols, nValCols, args, context);

            r = WriteRows(result, r, levels, subtotalDepth, totalsAtTop, nKeyCols, nValCols, args, depth: 1);

            if (totalsAtEnd && showTotals)
                WriteGrandTotal(result, r, levels, grandTotalStr, nKeyCols, nValCols, args, context);

            return result;
        }

        private int WriteRows(
    InMemoryRange result, int r,
    List<GroupLevel> levels,
    int subtotalDepth, bool subtotalsAtTop,
    int nKeyCols, int nValCols,
    GroupByBaseArgs args,
    int depth)
        {
            var functionHeaders = args.FunctionLayout == FunctionLayout.Vertical
                ? ResolveFunctionHeaders(args)
                : null;

            foreach (var level in levels)
            {
                bool writeSubtotal = subtotalDepth > 1 && depth <= subtotalDepth - 1;

                if (writeSubtotal && subtotalsAtTop)
                    r = WriteSubtotal(result, r, level, nKeyCols, nValCols, args);

                if (level.IsLeaf)
                {
                    foreach (var row in level.Rows)
                    {
                        if (args.FunctionLayout == FunctionLayout.Vertical)
                        {
                            for (int f = 0; f < args.Functions.Count; f++)
                            {
                                for (int c = 0; c < nKeyCols; c++)
                                    result.SetValue(r, c, row.KeyParts[c]);
                                result.SetValue(r, nKeyCols, functionHeaders[f]);
                                for (int c = 0; c < nValCols; c++)
                                    result.SetValue(r, nKeyCols + 1 + c, row.AggregatedValues[f][c]);
                                r++;
                            }
                        }
                        else
                        {
                            for (int c = 0; c < nKeyCols; c++)
                                result.SetValue(r, c, row.KeyParts[c]);
                            for (int f = 0; f < args.Functions.Count; f++)
                                for (int c = 0; c < nValCols; c++)
                                    result.SetValue(r, nKeyCols + f * nValCols + c, row.AggregatedValues[f][c]);
                            r++;
                        }
                    }
                }
                else
                {
                    r = WriteRows(result, r, level.Children, subtotalDepth, subtotalsAtTop, nKeyCols, nValCols, args, depth + 1);
                }

                if (writeSubtotal && !subtotalsAtTop)
                    r = WriteSubtotal(result, r, level, nKeyCols, nValCols, args);
            }
            return r;
        }


        private int CountDataRows(List<GroupLevel> levels)
        {
            int count = 0;
            foreach (var level in levels)
                count += level.IsLeaf
                    ? level.Rows.Count
                    : CountDataRows(level.Children);
            return count;
        }

        private int CountSubtotalRows(List<GroupLevel> levels, int subtotalDepth, int depth)
        {
            if (depth >= subtotalDepth) return 0;
            int count = levels.Count;
            foreach (var level in levels)
                if (!level.IsLeaf)
                    count += CountSubtotalRows(level.Children, subtotalDepth, depth + 1);
            return count;
        }

        private int WriteSubtotal(InMemoryRange result, int r, GroupLevel level, int nKeyCols, int nValCols, GroupByBaseArgs args)
        {
            var functionHeaders = ResolveFunctionHeaders(args);
            if (args.FunctionLayout == FunctionLayout.Vertical)
            {
                for (int f = 0; f < args.Functions.Count; f++)
                {
                    result.SetValue(r, 0, level.Key);
                    for (int c = 1; c < nKeyCols; c++)
                        result.SetValue(r, c, string.Empty);
                    result.SetValue(r, nKeyCols, functionHeaders[f]);
                    for (int c = 0; c < nValCols; c++)
                        result.SetValue(r, nKeyCols + 1 + c, level.SubtotalValues[f][c]);
                    r++;
                }
            }
            else
            {
                result.SetValue(r, 0, level.Key);
                for (int c = 1; c < nKeyCols; c++)
                    result.SetValue(r, c, string.Empty);
                for (int f = 0; f < args.Functions.Count; f++)
                    for (int c = 0; c < nValCols; c++)
                        result.SetValue(r, nKeyCols + f * nValCols + c, level.SubtotalValues[f][c]);
                r++;
            }
            return r;
        }

        private int WriteGrandTotal(InMemoryRange result, int r, List<GroupLevel> levels, string label, int nKeyCols, int nValCols, GroupByBaseArgs args, ParsingContext context)
        {
            var functionHeaders = ResolveFunctionHeaders(args);
            int nAllValCols = args.AllValuesInOrder.Count > 0 ? args.AllValuesInOrder[0].Length : 1;

            if (args.FunctionLayout == FunctionLayout.Vertical)
            {
                for (int f = 0; f < args.Functions.Count; f++)
                {
                    result.SetValue(r, 0, label);
                    for (int c = 1; c < nKeyCols; c++)
                        result.SetValue(r, c, string.Empty);
                    result.SetValue(r, nKeyCols, functionHeaders[f]);
                    for (int c = 0; c < nValCols; c++)
                    {
                        var colValues = args.AllValuesInOrder
                            .Select(v => new object[] { v[c] })
                            .ToList();
                        result.SetValue(r, nKeyCols + 1 + c, Aggregate(args.Functions[f], colValues, context,
                            args.Functions[f].EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null));
                    }
                    r++;
                }
            }
            else
            {
                result.SetValue(r, 0, label);
                for (int c = 1; c < nKeyCols; c++)
                    result.SetValue(r, c, string.Empty);
                for (int f = 0; f < args.Functions.Count; f++)
                    for (int c = 0; c < nValCols; c++)
                    {
                        var colValues = args.AllValuesInOrder
                            .Select(v => new object[] { v[c] })
                            .ToList();
                        result.SetValue(r, nKeyCols + f * nValCols + c, Aggregate(args.Functions[f], colValues, context,
                            args.Functions[f].EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null));
                    }
                r++;
            }
            return r;
        }

        /// <summary>Recursively collects all AggregatedValues from leaf GroupRows.</summary>
        private IEnumerable<object> CollectAggregatedValues(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows.Select(r => r.AggregatedValue);

            return level.Children.SelectMany(c => CollectAggregatedValues(c));
        }
    }
}
