using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class GroupbyFunctionBase : ExcelFunction
    {
        protected readonly LookupComparerBase _comparer = new SortByComparer();

        // -------------------------------------------------------
        // Enums & Constants
        // -------------------------------------------------------
        protected enum FieldHeaders
        {
            Missing = -1,
            No = 0,
            YesAndDontShow = 1,
            NoButGenerate = 2,
            YesAndShow = 3
        }

        protected const int TotalDepthNoTotals = 0;
        protected const int TotalDepthGrandOnly = 1;

        protected enum FieldRelationship
        {
            Hierarchy = 0,
            Table = 1
        }

        // -------------------------------------------------------
        // Shared argument container
        // -------------------------------------------------------
        protected class GroupByBaseArgs
        {
            public IRangeInfo RowFields { get; set; }
            public IRangeInfo Values { get; set; }
            public LambdaCalculator Function { get; set; }
            public FieldHeaders Headers { get; set; } = FieldHeaders.Missing;
            public int TotalDepth { get; set; } = 1;
            public int SortOrder { get; set; } = 1;
            public IRangeInfo FilterArray { get; set; } = null;
            public FieldRelationship FieldRelationship { get; set; } = FieldRelationship.Hierarchy;
            public List<object[]> AllValuesInOrder { get; set; } = new List<object[]>();
        }

        // -------------------------------------------------------
        // Shared data structures
        // -------------------------------------------------------
        protected class GroupLevel
        {
            public object Key { get; set; }
            public List<GroupLevel> Children { get; set; } = new List<GroupLevel>();
            public Dictionary<string, GroupLevel> ChildDict { get; set; } = null;
            public List<string> ChildOrder { get; set; } = null;
            public List<GroupRow> Rows { get; set; } = new List<GroupRow>();
            public object SubtotalValue { get; set; }
            public bool IsLeaf => Children.Count == 0;
        }

        protected class GroupRow
        {
            public object[] KeyParts { get; set; }
            public List<object[]> Values { get; set; } = new List<object[]>();
            public object AggregatedValue { get; set; }
        }

        // -------------------------------------------------------
        // Argument parsing (shared arguments 1-8)
        // -------------------------------------------------------
        protected bool TryParseBaseArgs(
            IList<FunctionArgument> arguments,
            out GroupByBaseArgs args,
            out CompileResult error)
        {
            args = new GroupByBaseArgs();
            error = null;

            if (!arguments[0].IsExcelRange)
                return Fail(eErrorType.Value, out error);
            args.RowFields = arguments[0].ValueAsRangeInfo;

            if (!arguments[1].IsExcelRange)
                return Fail(eErrorType.Value, out error);
            args.Values = arguments[1].ValueAsRangeInfo;

            if (args.RowFields.Size.NumberOfRows != args.Values.Size.NumberOfRows)
                return Fail(eErrorType.Value, out error);

            if (arguments[2].DataType != DataType.LambdaCalculation)
                return Fail(eErrorType.Value, out error);
            args.Function = arguments[2].Value as LambdaCalculator;
            if (args.Function == null)
                return Fail(eErrorType.Value, out error);

            // field_headers (optional)
            if (arguments.Count > 3 && arguments[3].Value != null)
            {
                var v = Convert.ToInt32(arguments[3].Value);
                if (!Enum.IsDefined(typeof(FieldHeaders), v))
                    return Fail(eErrorType.Value, out error);
                args.Headers = (FieldHeaders)v;
            }

            // total_depth (optional)
            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                var totalDepth = Convert.ToInt32(arguments[4].Value);
                if (Math.Abs(totalDepth) > args.RowFields.Size.NumberOfCols)
                    return Fail(eErrorType.Value, out error);
                args.TotalDepth = totalDepth;
            }

            // sort_order (optional)
            if (arguments.Count > 5 && arguments[5].Value != null)
                args.SortOrder = Convert.ToInt32(arguments[5].Value);

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

        protected bool Fail(eErrorType err, out CompileResult error)
        {
            error = CompileResult.GetErrorResult(err);
            return false;
        }

        // -------------------------------------------------------
        // Header resolution
        // -------------------------------------------------------
        protected FieldHeaders ResolveHeaders(GroupByBaseArgs args)
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

        // -------------------------------------------------------
        // Grouping
        // -------------------------------------------------------
        protected List<GroupLevel> BuildGroups(GroupByBaseArgs args, ParsingContext context)
        {
            var resolvedHeaders = ResolveHeaders(args);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
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

                    if (depth < nKeyCols - 1)
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
            AggregateTree(levels, args.Function, context);
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

        protected void AggregateTree(List<GroupLevel> levels, LambdaCalculator function, ParsingContext context)
        {
            foreach (var level in levels)
            {
                if (level.IsLeaf)
                {
                    foreach (var row in level.Rows)
                        row.AggregatedValue = Aggregate(function, row.Values, context);

                    var allVals = level.Rows.SelectMany(r => r.Values).ToList();
                    level.SubtotalValue = Aggregate(function, allVals, context);
                }
                else
                {
                    AggregateTree(level.Children, function, context);

                    var allVals = level.Children.SelectMany(c => GetAllValues(c)).ToList();
                    level.SubtotalValue = Aggregate(function, allVals, context);
                }
            }
        }

        protected List<object[]> GetAllValues(GroupLevel level)
        {
            if (level.IsLeaf)
                return level.Rows.SelectMany(r => r.Values).ToList();
            return level.Children.SelectMany(c => GetAllValues(c)).ToList();
        }

        protected object Aggregate(LambdaCalculator calculator, List<object[]> values, ParsingContext context)
        {
            int nRows = values.Count;
            int nCols = values.Count > 0 ? values[0].Length : 1;

            var range = new InMemoryRange(nRows, (short)nCols);
            for (int row = 0; row < nRows; row++)
                for (int col = 0; col < nCols; col++)
                    range.SetValue(row, col, values[row][col]);

            calculator.BeginCalculation();
            calculator.SetVariableValue(0, range, DataType.ExcelRange, context);
            return calculator.Execute(context).ResultValue;
        }

    }
}
