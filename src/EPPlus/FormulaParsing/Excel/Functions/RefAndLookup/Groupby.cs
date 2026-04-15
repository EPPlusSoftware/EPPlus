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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "8.6",
        Description = "Allows you to create a summary of your data via a formula. Supports grouping along one axis and aggregating the associated values.")]

    internal class GroupBy : GroupByFunctionBase
    {
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;
        public override int ArgumentMinLength => 3;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParseGroupByArgs(arguments, out var args, out var error))
                return error;
            var groups = BuildGroups(args, context);
            groups = ApplySort(groups, args);
            var result = BuildResult(groups, args, context);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private bool TryParseGroupByArgs(IList<FunctionArgument> arguments,
            out GroupByBaseArgs args,
            out CompileResult error)
        {
            args = new GroupByBaseArgs();
            error = null;

            if (!arguments[0].IsExcelRange) // TODO. Man kan skicka in enskilda celler i rowfields och values, så detta är fel.
                return Fail(eErrorType.Value, out error);
            args.RowFields = arguments[0].ValueAsRangeInfo;

            if (!arguments[1].IsExcelRange)
                return Fail(eErrorType.Value, out error);
            args.Values = arguments[1].ValueAsRangeInfo;

            if (args.RowFields.Size.NumberOfRows != args.Values.Size.NumberOfRows)
                return Fail(eErrorType.Value, out error);

            if (!TryParseFunctionArg(arguments[2], args.Functions, out LambdaCalculator function, out FunctionLayout layout))
                return Fail(eErrorType.Value, out error);

            args.Function = function;
            args.FunctionLayout = layout;

            if (args.Functions.Count == 0)
                return Fail(eErrorType.Value, out error);

            // field_headers (optional)
            if (arguments.Count > 3 && arguments[3].Value != null)
            {
                var v = Convert.ToInt32(arguments[3].Value);
                if (!Enum.IsDefined(typeof(FieldHeaders), v))
                    return Fail(eErrorType.Value, out error);
                args.Headers = (FieldHeaders)v;
            }
            else if (args.Functions.Count > 1) // In excel, if multiple functions are included, headers are by default displayed.
            {
                args.Headers = FieldHeaders.YesAndShow;
            }

            // total_depth (optional)
            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[4], args.RowFields.Size.NumberOfCols, out int totalDepth))
                    return Fail(eErrorType.Value, out error);
                args.TotalDepth = totalDepth;
            }

            // sort_order (optional)
            if (arguments.Count > 5 && arguments[5].Value != null)
            {                
                args.SortOrders = ParseSortOrderArg(arguments[5]);
            }

            // filter_array (optional)
            if (arguments.Count > 6 && arguments[6].IsExcelRange)
                args.FilterArray = arguments[6].ValueAsRangeInfo;

            // field_relationship (optional)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                var v = Convert.ToInt32(arguments[7].Value);
                if (!Enum.IsDefined(typeof(FieldRelationship), v))
                    return Fail(eErrorType.Value, out error);
                if (v == (int)FieldRelationship.Table && Math.Abs(args.TotalDepth) > 1)
                    return Fail(eErrorType.Value, out error);
                args.FieldRelationship = (FieldRelationship)v;
            }

            return true;
        }
        
        // -------------------------------------------------------
        // Build result
        // -------------------------------------------------------        

        private InMemoryRange BuildResult(List<GroupLevel> levels, GroupByBaseArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
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
