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
        private class GroupbyArgs
        {
            public IRangeInfo RowFields { get; set; }
            public IRangeInfo Values { get; set; }
            public LambdaCalculator Function { get; set; }
            public FieldHeaders Headers { get; set; } = FieldHeaders.Missing; // Default
            public TotalDepth TotalDepth { get; set; } = TotalDepth.GrandTotals; // Default
            public int SortOrder { get; set; } = 1; // Default is first column ascending.
            public IRangeInfo FilterArray { get; set; } = null;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParseArgs(arguments, context, out var args, out var error))
                return error;
            var groups = BuildGroups(args, context);
            groups = ApplySort(groups, args.SortOrder);
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
                if (!Enum.IsDefined(typeof(TotalDepth), Convert.ToInt32(arguments[4].Value)))
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

            return true;
        }
        /// <summary>Represents one group key with its collected values.</summary>
        private class GroupRow
        {
            public object Key { get; set; }
            public List<object> Values { get; set; } = new List<object>();
            public object AggregatedValue { get; set; }
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

        private List<GroupRow> BuildGroups(GroupbyArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            int startRow = hasHeaders ? 1 : 0;

            var dict = new Dictionary<object, GroupRow>();
            var order = new List<object>();

            for (int r = startRow; r < args.RowFields.Size.NumberOfRows; r++)
            {
                // Apply filter_array if present
                if (args.FilterArray != null)
                {
                    var filterVal = args.FilterArray.GetValue(r, 0);
                    if (filterVal is bool b && !b) continue;
                    if (filterVal is int i && i == 0) continue;
                }

                //var key = args.RowFields.GetValue(r, 0)?.ToString() ?? string.Empty;
                //var val = args.Values.GetValue(r, 0);

                var key = args.RowFields.GetOffset(r, 0);
                    //?.ToString() ?? string.Empty;
                var val = args.Values.GetOffset(r, 0);

                if (!dict.TryGetValue(key, out var group))
                {
                    group = new GroupRow { Key = key };
                    dict[key] = group;
                    order.Add(key);
                }
                group.Values.Add(val);
            }

            // Aggregate each group using the lambda
            var groups = order.Select(k => dict[k]).ToList();
            foreach (var g in groups)
                g.AggregatedValue = Aggregate(args.Function, g.Values, context);

            return groups;
        }

        /// <summary>
        /// Aggregates a group's values by passing them as an in-memory range
        /// to the LambdaCalculator, mirroring the pattern used in Map.
        /// </summary>
        private object Aggregate(LambdaCalculator calculator, List<object> values, ParsingContext context)
        {
            // Build a single-column in-memory range from the group's values
            var range = new InMemoryRange(values.Count, 1);
            for (int i = 0; i < values.Count; i++)
                range.SetValue(i, 0, values[i]);

            calculator.BeginCalculation();
            // Pass the range as the single variable (e.g. the x in LAMBDA(x, SUM(x)))
            calculator.SetVariableValue(0, range, DataType.ExcelRange, context);
            var result = calculator.Execute(context);
            return result.ResultValue;
        }

        // -------------------------------------------------------
        // Sorting
        // -------------------------------------------------------
        private List<GroupRow> ApplySort(List<GroupRow> groups, int sortOrder)
        {
            //if (sortOrder == 0) sortOrder = 1;

            bool desc = sortOrder < 0;
            int col = Math.Abs(sortOrder);

            if (col == 1)
            {
                return desc ? groups.OrderByDescending(g => g.Key, _comparer).ToList()
                            : groups.OrderBy(g => g.Key, _comparer).ToList();
            }

            // Column 2+ sorts on aggregated value - handle both numeric and text
            return desc
                ? groups.OrderByDescending(g => g.AggregatedValue as IComparable, _comparer).ToList()
                : groups.OrderBy(g => g.AggregatedValue as IComparable, _comparer).ToList();
        }

        // -------------------------------------------------------
        // Build result
        // -------------------------------------------------------
        private InMemoryRange BuildResult(List<GroupRow> groups, GroupbyArgs args)
        {
            bool showHeaders = args.Headers == FieldHeaders.YesAndShow
                            || args.Headers == FieldHeaders.NoButGenerate;
            bool totalsAtEnd = args.TotalDepth == TotalDepth.GrandTotals
                            || args.TotalDepth == TotalDepth.GrandAndSubtotals;
            bool totalsAtTop = args.TotalDepth == TotalDepth.GrandTotalsAtTop
                            || args.TotalDepth == TotalDepth.GrandAndSubtotalsAtTop;
            bool showTotals = args.TotalDepth != TotalDepth.NoTotals;

            int rowCount = groups.Count
                + (showHeaders ? 1 : 0)
                + (showTotals ? 1 : 0);

            var result = new InMemoryRange(rowCount, 2);
            int r = 0;

            if (showHeaders)
            {
                result.SetValue(r, 0, args.Headers == FieldHeaders.NoButGenerate
                    ? "Field 1" : args.RowFields.GetValue(0, 0)?.ToString());
                result.SetValue(r, 1, args.Headers == FieldHeaders.NoButGenerate
                    ? "Field 2" : args.Values.GetValue(0, 0)?.ToString());
                r++;
            }

            if (totalsAtTop && showTotals)
            {
                result.SetValue(r, 0, "Total");
                result.SetValue(r, 1, SumAggregated(groups));
                r++;
            }

            foreach (var g in groups)
            {
                result.SetValue(r, 0, g.Key);
                result.SetValue(r, 1, g.AggregatedValue);
                r++;
            }

            if (totalsAtEnd && showTotals)
            {
                result.SetValue(r, 0, "Total");
                result.SetValue(r, 1, SumAggregated(groups));
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
