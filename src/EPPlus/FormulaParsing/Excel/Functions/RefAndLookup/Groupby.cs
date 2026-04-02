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
            if (args.SortOrder == 0) return levels;

            bool desc = args.SortOrder < 0;
            int col = Math.Abs(args.SortOrder);
            bool sortOnAggregated = col > args.RowFields.Size.NumberOfCols;

            // desc only applies at the depth that matches the sort column
            bool descThisLevel = desc && depth == col;

            if (args.FieldRelationship == FieldRelationship.Table)
            {
                // Collect all leaf rows recursively
                var allRows = levels.SelectMany(l => CollectLeafRows(l)).ToList();
                allRows = SortRows(allRows, col, desc, sortOnAggregated);

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
                levels = sortOnAggregated
                    ? (descThisLevel ? levels.OrderByDescending(l => l.SubtotalValue as IComparable, _comparer).ToList()
                                     : levels.OrderBy(l => l.SubtotalValue as IComparable, _comparer).ToList())
                    : (descThisLevel ? levels.OrderByDescending(l => l.Key as IComparable, _comparer).ToList()
                                     : levels.OrderBy(l => l.Key as IComparable, _comparer).ToList());

                foreach (var level in levels)
                {
                    if (!level.IsLeaf)
                        level.Children = ApplySort(level.Children, args, depth + 1);
                    else
                        level.Rows = SortRows(level.Rows, depth, depth == col ? desc : false, sortOnAggregated);
                }

                return levels;
            }
        }

        private IEnumerable<GroupRow> CollectLeafRows(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows;
            return level.Children.SelectMany(c => CollectLeafRows(c));
        }

        private List<GroupRow> SortRows(List<GroupRow> rows, int col, bool desc, bool sortOnAggregated)
        {
            if (rows == null || rows.Count == 0) return rows;

            if (sortOnAggregated)
            {
                return desc ? rows.OrderByDescending(r => r.AggregatedValue as IComparable, _comparer).ToList()
                            : rows.OrderBy(r => r.AggregatedValue as IComparable, _comparer).ToList();
            }

            // col is 1-based, KeyParts is 0-based
            int keyIndex = Math.Min(col - 1, rows[0].KeyParts.Length - 1);
            return desc ? rows.OrderByDescending(r => r.KeyParts[keyIndex] as IComparable, _comparer).ToList()
                        : rows.OrderBy(r => r.KeyParts[keyIndex] as IComparable, _comparer).ToList();
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

            var result = new InMemoryRange(totalRows, (short)nCols); // denna är ett för lite. TODO
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
                                    result.SetValue(r, nKeyCols + 1 + c, row.AggregatedValues[f]);
                                r++;
                            }
                        }
                        else
                        {
                            for (int c = 0; c < nKeyCols; c++)
                                result.SetValue(r, c, row.KeyParts[c]);
                            for (int f = 0; f < args.Functions.Count; f++)
                                for (int c = 0; c < nValCols; c++)
                                    result.SetValue(r, nKeyCols + f * nValCols + c, row.AggregatedValues[f]);
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
                        result.SetValue(r, nKeyCols + 1 + c, level.SubtotalValues[f]);
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
                        result.SetValue(r, nKeyCols + f * nValCols + c, level.SubtotalValues[f]);
                r++;
            }
            return r;
        }

        private int WriteGrandTotal(InMemoryRange result, int r, List<GroupLevel> levels, string label, int nKeyCols, int nValCols, GroupByBaseArgs args, ParsingContext context)
        {
            var functionHeaders = ResolveFunctionHeaders(args);
            if (args.FunctionLayout == FunctionLayout.Vertical)
            {
                for (int f = 0; f < args.Functions.Count; f++)
                {
                    result.SetValue(r, 0, label);
                    for (int c = 1; c < nKeyCols; c++)
                        result.SetValue(r, c, string.Empty);
                    result.SetValue(r, nKeyCols, functionHeaders[f]);
                    result.SetValue(r, nKeyCols + 1, Aggregate(args.Functions[f], args.AllValuesInOrder, context,
                        args.Functions[f].EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null));
                    r++;
                }
            }
            else
            {
                result.SetValue(r, 0, label);
                for (int c = 1; c < nKeyCols; c++)
                    result.SetValue(r, c, string.Empty);
                for (int f = 0; f < args.Functions.Count; f++)
                    result.SetValue(r, nKeyCols + f * nValCols, Aggregate(args.Functions[f], args.AllValuesInOrder, context,
                        args.Functions[f].EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null));
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
