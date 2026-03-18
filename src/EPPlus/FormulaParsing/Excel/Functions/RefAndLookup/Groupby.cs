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

    internal class Groupby : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;
        public override int ArgumentMinLength => 3;
        private readonly LookupComparerBase _comparer = new SortByComparer();
        private enum FieldHeaders
        {
            Missing = -1, // Default 
            No = 0,
            YesAndDontShow = 1,
            NoButGenerate = 2,
            YesAndShow = 3
        }
        private enum TotalDepth
        {
            Missing = -3,
            GrandAndSubtotalsAtTop = -2,
            GrandTotalsAtTop = -1,            
            NoTotals = 0,
            GrandTotals = 1,
            GrandAndSubtotals = 2,
        }
        private enum FieldRelationship
        {
            Hierarchy = 0,
            Table = 1
        }
        private class GroupbyArgs
        {
            public IRangeInfo RowFields { get; set; }
            public IRangeInfo Values { get; set; }
            public LambdaCalculator Function { get; set; }
            public FieldHeaders Headers { get; set; } = FieldHeaders.Missing; // Default
            public TotalDepth TotalDepth { get; set; } = TotalDepth.GrandTotals; // Default
            public int SortOrder { get; set; } = 1; // Default is first column ascending.
            public IRangeInfo FilterArray { get; set; } = null;
            public FieldRelationship FieldRelationship { get; set; } = FieldRelationship.Hierarchy;
        }

        /// <summary>Represents one level, with the topkey that represents the key in the first column. </summary>
        private class GroupLevel
        {
            public object TopKey { get; set; } // First column value
            public List<GroupRow> Rows { get; set; } = new List<GroupRow>();
            public object SubtotalValue { get; set; }
        }

        /// <summary>Represents one group key with its collected values.</summary>
        private class GroupRow
        {
            public string Key { get; set; }
            public object[] KeyParts { get; set; }
            public List<object[]> Values { get; set; } = new List<object[]>();
            public object AggregatedValue { get; set; }
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParseArgs(arguments, context, out var args, out var error))
                return error;
            var groups = BuildGroups(args, context);
            groups = ApplySort(groups, args, args.SortOrder);
            var result = BuildResult(groups, args);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private bool TryParseArgs(
            IList<FunctionArgument> arguments,
            ParsingContext context,
            out GroupbyArgs args,
            out CompileResult error)
        {
            args = new GroupbyArgs();
            error = null;

            if (!arguments[0].IsExcelRange)
            {
                error = CompileResult.GetErrorResult(eErrorType.Value);
                return false;
            }
            args.RowFields = arguments[0].ValueAsRangeInfo;

            if (!arguments[1].IsExcelRange)
            {
                error = CompileResult.GetErrorResult(eErrorType.Value);
                return false;
            }
            args.Values = arguments[1].ValueAsRangeInfo;

            // Validate that row_fields and values have the same number of rows
            if (args.RowFields.Size.NumberOfRows != args.Values.Size.NumberOfRows)
            {
                error = CompileResult.GetErrorResult(eErrorType.Value);
                return false;
            }

            // function (required) – resolve the aggregation function by name
            if (arguments[2].DataType != DataType.LambdaCalculation)
            {
                error = CompileResult.GetErrorResult(eErrorType.Value);
                return false;
            }
            var calculator = arguments[2].Value as LambdaCalculator;
            args.Function = calculator;            if (args.Function == null)
            {
                error = CompileResult.GetErrorResult(eErrorType.Value);
                return false;
            }

            // field_headers (optional, default = HasHeadersAndShow)
            if (arguments.Count > 3 && arguments[3].Value != null)
            {
                if (!Enum.IsDefined(typeof(FieldHeaders), Convert.ToInt32(arguments[3].Value)))
                {
                    error = CompileResult.GetErrorResult(eErrorType.Value);
                    return false;
                }
                args.Headers = (FieldHeaders)Convert.ToInt32(arguments[3].Value);
            }

            // total_depth (optional, default = GrandTotals)
            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                if (!Enum.IsDefined(typeof(TotalDepth), Convert.ToInt32(arguments[4].Value)) 
                    || args.RowFields.Size.NumberOfCols > Convert.ToInt32(arguments[4].Value)
                    || args.RowFields.Size.NumberOfCols * -1 < Convert.ToInt32(arguments[4].Value) * -1 )
                {
                    error = CompileResult.GetErrorResult(eErrorType.Value);
                    return false;
                }
                args.TotalDepth = (TotalDepth)Convert.ToInt32(arguments[4].Value);
            }

            // sort_order (optional, default = 0 / no sort)
            if (arguments.Count > 5 && arguments[5].Value != null)
                args.SortOrder = Convert.ToInt32(arguments[5].Value);

            // filter_array (optional)
            if (arguments.Count > 6 && arguments[6].IsExcelRange)
                args.FilterArray = arguments[6].ValueAsRangeInfo;

            // field_relationship (optional)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                if (!Enum.IsDefined(typeof(FieldRelationship), Convert.ToInt32(arguments[7].Value)) 
                    || args.TotalDepth == TotalDepth.GrandAndSubtotals || args.TotalDepth == TotalDepth.GrandAndSubtotalsAtTop)
                {
                    error = CompileResult.GetErrorResult(eErrorType.Value);
                    return false;
                }
                args.FieldRelationship = (FieldRelationship)Convert.ToInt32(arguments[7].Value);
            }

            return true;
        }        

        private FieldHeaders ResolveHeaders (GroupbyArgs args)
        {
            if (args.Headers != FieldHeaders.Missing)
                return args.Headers;

            if (args.Values.Size.NumberOfRows < 2)
                return FieldHeaders.No;

            var first = args.Values.GetValue(0, 0);
            var second = args.Values.GetValue(1, 0);

            bool firstIsText = first is string;
            bool secondIsNumber = second is double || second is int || second is long || second is float;

            return firstIsText && secondIsNumber
                ? FieldHeaders.YesAndDontShow
                : FieldHeaders.No;
        }

        private List<GroupLevel> BuildGroups(GroupbyArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            int startRow = hasHeaders ? 1 : 0;

            var levelDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
            var levelOrder = new List<string>();
            var rowDict = new Dictionary<string, GroupRow>(StringComparer.OrdinalIgnoreCase);

            for (int r = startRow; r < args.RowFields.Size.NumberOfRows; r++)
            {
                // Apply filter_array if present
                if (args.FilterArray != null)
                {
                    var filterVal = args.FilterArray.GetOffset(r, 0);
                    if (filterVal is bool b && !b) continue;
                    if (filterVal is int i && i == 0) continue;
                }

                // Build composite key from all columns in RowFields
                int nKeyCols = args.RowFields.Size.NumberOfCols;
                var keyParts = new object[nKeyCols];
                for (int c = 0; c < nKeyCols; c++)
                    keyParts[c] = args.RowFields.GetOffset(r, c);

                var keyStrings = new string[keyParts.Length];
                for (int c = 0; c < keyParts.Length; c++)
                    keyStrings[c] = keyParts[c]?.ToString() ?? string.Empty;
                var key = string.Join("|", keyStrings);

                // Top-level key is always the first column
                var topKeyStr = keyStrings[0];

                // Collect all columns from values range for this row
                int nCols = args.Values.Size.NumberOfCols;
                var rowVals = new object[nCols];
                for (int c = 0; c < nCols; c++)
                    rowVals[c] = args.Values.GetOffset(r, c);

                // Create GroupLevel if needed
                if (!levelDict.TryGetValue(topKeyStr, out var level))
                {
                    level = new GroupLevel { TopKey = keyParts[0] };
                    levelDict[topKeyStr] = level;
                    levelOrder.Add(topKeyStr);
                }

                // Create GroupRow if needed
                if (!rowDict.TryGetValue(key, out var group))
                {
                    group = new GroupRow { Key = key, KeyParts = keyParts };
                    rowDict[key] = group;
                    level.Rows.Add(group);
                }
                group.Values.Add(rowVals);
            }

            // Aggregate each GroupRow and each GroupLevel's subtotal
            var levels = levelOrder.Select(k => levelDict[k]).ToList();
            foreach (var level in levels)
            {
                foreach (var row in level.Rows)
                    row.AggregatedValue = Aggregate(args.Function, row.Values, context);

                // Subtotal = aggregate of all values in this level
                var allLevelValues = level.Rows.SelectMany(r => r.Values).ToList();
                level.SubtotalValue = Aggregate(args.Function, allLevelValues, context);
            }

            return levels;
        }

        /// <summary>
        /// Aggregates a group's values by passing them as an in-memory range
        /// to the LambdaCalculator, mirroring the pattern used in Map.
        /// </summary>
        private object Aggregate(LambdaCalculator calculator, List<object[]> values, ParsingContext context)
        {
            int nRows = values.Count;
            int nCols = values.Count > 0 ? values[0].Length : 1;

            var range = new InMemoryRange(nRows, (short)nCols);
            for (int row = 0; row < nRows; row++)
                for (int col = 0; col < nCols; col++)
                    range.SetValue(row, col, values[row][col]);

            calculator.BeginCalculation();
            calculator.SetVariableValue(0, range, DataType.ExcelRange, context);
            var result = calculator.Execute(context);
            return result.ResultValue;
        }

        // -------------------------------------------------------
        // Sorting
        // -------------------------------------------------------
        private List<GroupLevel> ApplySort(List<GroupLevel> levels, GroupbyArgs args, int sortOrder)
        {
            if (sortOrder == 0) return levels;

            bool desc = sortOrder < 0;
            int col = Math.Abs(sortOrder);
            bool sortOnAggregated = col > args.RowFields.Size.NumberOfRows;

            //// Sort levels by TopKey or subtotal
            //var sortedLevels = col == 1
            //    ? (desc ? levels.OrderByDescending(l => l.TopKey as IComparable, _comparer).ToList()
            //            : levels.OrderBy(l => l.TopKey as IComparable, _comparer).ToList())
            //    : (desc ? levels.OrderByDescending(l => l.SubtotalValue as IComparable, _comparer).ToList()
            //            : levels.OrderBy(l => l.SubtotalValue as IComparable, _comparer).ToList());

            //// Sort rows within each level by their KeyParts or aggregated value
            //foreach (var level in sortedLevels)
            //{
            //    level.Rows = col == 1
            //        ? (desc ? level.Rows.OrderByDescending(r => r.KeyParts[0] as IComparable, _comparer).ToList()
            //                : level.Rows.OrderBy(r => r.KeyParts[0] as IComparable, _comparer).ToList())
            //        : (desc ? level.Rows.OrderByDescending(r => r.AggregatedValue as IComparable, _comparer).ToList()
            //                : level.Rows.OrderBy(r => r.AggregatedValue as IComparable, _comparer).ToList());
            //}

            //return sortedLevels;
            if (args.FieldRelationship == FieldRelationship.Table)
            {
                // Table: sort all GroupRows globally and independently, rebuild levels
                var allRows = levels.SelectMany(l => l.Rows).ToList();
                allRows = SortRows(allRows, col, desc, sortOnAggregated);

                // Rebuild levels in the new row order
                var newLevelDict = new Dictionary<string, GroupLevel>();
                var newLevelOrder = new List<string>();
                foreach (var row in allRows)
                {
                    var topKey = row.KeyParts[0]?.ToString() ?? string.Empty;
                    if (!newLevelDict.TryGetValue(topKey, out var level))
                    {
                        level = new GroupLevel { TopKey = row.KeyParts[0] };
                        newLevelDict[topKey] = level;
                        newLevelOrder.Add(topKey);
                    }
                    level.Rows.Add(row);
                }
                return newLevelOrder.Select(k => newLevelDict[k]).ToList();
            }
            else
            {
                // Hierarchy: sort levels on TopKey or aggregated, then sort rows within each level
                levels = sortOnAggregated
                    ? (desc ? levels.OrderByDescending(l => l.SubtotalValue as IComparable).ToList()
                            : levels.OrderBy(l => l.SubtotalValue as IComparable).ToList())
                    : (desc ? levels.OrderByDescending(l => l.TopKey as IComparable).ToList()
                            : levels.OrderBy(l => l.TopKey as IComparable).ToList());

                // Sort rows within each level
                foreach (var level in levels)
                    level.Rows = SortRows(level.Rows, col, desc, sortOnAggregated);

                return levels;
            }
        }

        private List<GroupRow> SortRows(List<GroupRow> rows, int col, bool desc, bool sortOnAggregated)
        {
            if (sortOnAggregated)
            {
                return desc ? rows.OrderByDescending(r => r.AggregatedValue as IComparable).ToList()
                            : rows.OrderBy(r => r.AggregatedValue as IComparable).ToList();
            }

            // col is 1-based, KeyParts is 0-based
            int keyIndex = Math.Min(col - 1, rows[0].KeyParts.Length - 1);
            return desc ? rows.OrderByDescending(r => r.KeyParts[keyIndex] as IComparable).ToList()
                        : rows.OrderBy(r => r.KeyParts[keyIndex] as IComparable).ToList();
        }
        // -------------------------------------------------------
        // Build result
        // -------------------------------------------------------        

        private InMemoryRange BuildResult(List<GroupLevel> levels, GroupbyArgs args)
        {
            var resolvedHeaders = ResolveHeaders(args);
            bool showHeaders = args.Headers == FieldHeaders.YesAndShow
                             || args.Headers == FieldHeaders.NoButGenerate;
            bool totalsAtEnd = args.TotalDepth == TotalDepth.GrandTotals
                             || args.TotalDepth == TotalDepth.GrandAndSubtotals;
            bool totalsAtTop = args.TotalDepth == TotalDepth.GrandTotalsAtTop
                             || args.TotalDepth == TotalDepth.GrandAndSubtotalsAtTop;
            bool showTotals = args.TotalDepth != TotalDepth.NoTotals;
            bool showSubtotals = args.TotalDepth == TotalDepth.GrandAndSubtotals
                              || args.TotalDepth == TotalDepth.GrandAndSubtotalsAtTop;

            int nKeyCols = args.RowFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;
            int nCols = nKeyCols + nValCols;

            // Calculate total number of rows needed
            int dataRows = levels.Sum(l => l.Rows.Count);
            int subtotalRows = showSubtotals ? levels.Count : 0;
            int totalRows = dataRows + subtotalRows
                             + (showHeaders ? 1 : 0)
                             + (showTotals ? 1 : 0);

            var result = new InMemoryRange(totalRows, (short)nCols);
            int r = 0;

            // Header row
            if (showHeaders)
            {
                for (int c = 0; c < nKeyCols; c++)
                    result.SetValue(r, c, resolvedHeaders == FieldHeaders.NoButGenerate
                        ? $"Field {c + 1}"
                        : args.RowFields.GetOffset(0, c)?.ToString());
                for (int c = 0; c < nValCols; c++)
                    result.SetValue(r, nKeyCols + c, resolvedHeaders == FieldHeaders.NoButGenerate
                        ? $"Field {nKeyCols + c + 1}"
                        : args.Values.GetOffset(0, c)?.ToString());
                r++;
            }

            string totalString;
            if (args.TotalDepth == TotalDepth.GrandAndSubtotalsAtTop || args.TotalDepth == TotalDepth.GrandAndSubtotals)
            {
                totalString = "Grand Total";
            }
            else
            {
                totalString = "Total";
            }
            // Grand total at top
            if (totalsAtTop && showTotals)
            {
                result.SetValue(r, 0, totalString);
                result.SetValue(r, nKeyCols, levels.SelectMany(l => l.Rows).Sum(row => Convert.ToDouble(row.AggregatedValue)));
                r++;
            }

            // Data rows and subtotals
            foreach (var level in levels)
            {
                foreach (var row in level.Rows)
                {
                    for (int c = 0; c < nKeyCols; c++)
                        result.SetValue(r, c, row.KeyParts[c]);
                    result.SetValue(r, nKeyCols, row.AggregatedValue);
                    r++;
                }

                if (showSubtotals)
                {
                    result.SetValue(r, 0, level.TopKey);
                    result.SetValue(r, nKeyCols, level.SubtotalValue);
                    r++;
                }
            }

            // Grand total at bottom
            if (totalsAtEnd && showTotals)
            {
                result.SetValue(r, 0, totalString);
                result.SetValue(r, nKeyCols, levels.SelectMany(l => l.Rows).Sum(row => Convert.ToDouble(row.AggregatedValue)));
            }

            return result;
        }

        /// <summary>Sums aggregated values where they are numeric.</summary>
        private object SumAggregated(List<GroupRow> groups)
        {
            var nums = groups
                .Select(g => g.AggregatedValue)
                .Where(v => v is IConvertible && !(v is string))
                .Select(v => Convert.ToDouble(v))
                .ToList();

            return nums.Any() ? (object)nums.Sum() : string.Empty;
        }
    }
}
