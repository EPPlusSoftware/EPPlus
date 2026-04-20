/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/4/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.LookupAndReference,
       EPPlusVersion = "8.6",
       Description = "Allows you to create a summary of your data via a formula. It supports grouping along two axis and aggregating the associated values.")]
    internal class PivotBy : GroupByFunctionBase
    {
        public override int ArgumentMinLength => 3; // kanske ska vara 4
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParsePivotByArgs(arguments, out var args, out var error))
                return error;
            BuildPivotData(args, context, out var rowLeaves, out var colLeaves, out var pivotMap);
            rowLeaves = ApplyRowSort(rowLeaves, args);
            colLeaves = ApplyColSort(colLeaves, args);
            var result = RenderPivot(rowLeaves, colLeaves, pivotMap, args, context);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);            
        }

        // 6. Räkna ut dimensioner: nRows, nCols
        // 7. Skriv kolumnhuvuden (ett pass)
        // 8. Skriv datarader rad för rad, slå upp pivotMap[(rowKey, colKey)] per kolumn
        // 9. Skriv grand total-rad/-kolumn

        protected bool TryParsePivotByArgs(IList<FunctionArgument> arguments,
            out PivotByArgs args,
            out CompileResult error)
        {
            args = new PivotByArgs();
            error = null;

            args.RowFields = arguments[0].ValueAsRangeInfo;
            args.ColFields = arguments[1].ValueAsRangeInfo;
            args.Values = arguments[2].ValueAsRangeInfo;

            if (args.RowFields.Size.NumberOfRows != args.Values.Size.NumberOfRows)
                return Fail(eErrorType.Value, out error);

            if (!TryParseFunctionArg(arguments[3], args.Functions, out LambdaCalculator function, out FunctionLayout layout))
                return Fail(eErrorType.Value, out error);

            args.Function = function;
            args.FunctionLayout = layout;

            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                var v = Convert.ToInt32(arguments[3].Value);
                if (!Enum.IsDefined(typeof(FieldHeaders), v))
                    return Fail(eErrorType.Value, out error);
                args.Headers = (FieldHeaders)v;
            }

            // Total depth for rows (optional)
            if (arguments.Count > 5 && arguments[5].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[5], args.RowFields.Size.NumberOfCols, out int rowTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.RowTotalDepth = rowTotalDepth;
            }

            // SortOrder for RowFields (optional)
            if (arguments.Count > 6 && arguments[6].Value != null)
                args.RowSortOrders = ParseSortOrderArg(arguments[6]);

            // TotalDepth for columns (optional)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[7], args.ColFields.Size.NumberOfCols, out int colTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.ColTotalDepth = colTotalDepth;
            }

            // SortOrder for ColFields (optional)
            if (arguments.Count > 8 && arguments[8].Value != null)
                args.ColSortOrders = ParseSortOrderArg(arguments[8]);

            if (arguments.Count > 9 && arguments[9].IsExcelRange)
                args.FilterArray = arguments[9].ValueAsRangeInfo;

            // RelativeTo (optional)
            if (arguments.Count > 10 && arguments[10].Value != null)
            {
                var v = Convert.ToInt32(arguments[10].Value);
                if (!Enum.IsDefined(typeof(RelativeTo), v))
                    return Fail(eErrorType.Value, out error);
                args.RelativeTo = (RelativeTo)v;
            }

            return true;
        }


        private void BuildPivotData(
      PivotByArgs args,
      ParsingContext context,
      out List<LeafWithPath> rowLeaves,
      out List<LeafWithPath> colLeaves,
      out Dictionary<string, Dictionary<string, List<object[]>>> pivotMap)
        {
            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            int startRow = hasHeaders ? 1 : 0;

            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColKeyCols = args.ColFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;

            var rowLeafDict = new Dictionary<string, LeafWithPath>(StringComparer.OrdinalIgnoreCase);
            var rowLeafOrder = new List<string>();
            var colLeafDict = new Dictionary<string, LeafWithPath>(StringComparer.OrdinalIgnoreCase);
            var colLeafOrder = new List<string>();
            pivotMap = new Dictionary<string, Dictionary<string, List<object[]>>>(StringComparer.OrdinalIgnoreCase);

            int nRows = args.RowFields.Size.NumberOfRows;

            for (int r = startRow; r < nRows; r++)
            {
                if (args.FilterArray != null)
                {
                    var fv = args.FilterArray.GetOffset(r, 0);
                    if (fv is bool b && !b) continue;
                    if (fv is int i && i == 0) continue;
                }

                var rowKeyParts = new object[nRowKeyCols];
                for (int c = 0; c < nRowKeyCols; c++)
                    rowKeyParts[c] = args.RowFields.GetOffset(r, c);

                var colKeyParts = new object[nColKeyCols];
                for (int c = 0; c < nColKeyCols; c++)
                    colKeyParts[c] = args.ColFields.GetOffset(r, c);

                var vals = new object[nValCols];
                for (int c = 0; c < nValCols; c++)
                    vals[c] = args.Values.GetOffset(r, c);

                // Radlöv
                string rowKey = MakePivotKey(rowKeyParts);
                if (!rowLeafDict.ContainsKey(rowKey))
                {
                    var leaf = new GroupLevel { Key = rowKeyParts[rowKeyParts.Length - 1] };
                    rowLeafDict[rowKey] = new LeafWithPath(leaf, rowKeyParts);
                    rowLeafOrder.Add(rowKey);
                }
                var rowLeafEntry = rowLeafDict[rowKey];
                var existingRow = rowLeafEntry.Leaf.Rows.FirstOrDefault(rw => MakePivotKey(rw.KeyParts) == rowKey);
                if (existingRow == null)
                {
                    existingRow = new GroupRow { KeyParts = rowKeyParts };
                    rowLeafEntry.Leaf.Rows.Add(existingRow);
                }
                existingRow.Values.Add(vals);

                // Kolumnlöv
                string colKey = MakePivotKey(colKeyParts);
                if (!colLeafDict.ContainsKey(colKey))
                {
                    var leaf = new GroupLevel { Key = colKeyParts[colKeyParts.Length - 1] };
                    colLeafDict[colKey] = new LeafWithPath(leaf, colKeyParts);
                    colLeafOrder.Add(colKey);
                }

                // Pivotkartan
                if (!pivotMap.TryGetValue(rowKey, out var colMap))
                {
                    colMap = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
                    pivotMap[rowKey] = colMap;
                }
                if (!colMap.TryGetValue(colKey, out var cellVals))
                {
                    cellVals = new List<object[]>();
                    colMap[colKey] = cellVals;
                }
                cellVals.Add(vals);
                args.AllValuesInOrder.Add(vals);
            }

            rowLeaves = rowLeafOrder.Select(k => rowLeafDict[k]).ToList();
            colLeaves = colLeafOrder.Select(k => colLeafDict[k]).ToList();

            // Aggregera SubtotalValues per radlöv (används för rad-totals)
            foreach (var rl in rowLeaves)
                AggregateLeaf(rl.Leaf, args, context);
        }

        private void AggregateLeaf(GroupLevel leaf, PivotByArgs args, ParsingContext context)
        {
            var allVals = leaf.Rows.SelectMany(r => r.Values).ToList();
            leaf.SubtotalValues = args.Functions.Select(f =>
            {
                int nValCols = allVals[0].Length;
                var result = new object[nValCols];
                for (int col = 0; col < nValCols; col++)
                {
                    var colValues = allVals.Select(v => new object[] { v[col] }).ToList();
                    result[col] = Aggregate(f, colValues, context,
                        f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                }
                return result;
            }).ToList();
            leaf.SubtotalValue = leaf.SubtotalValues[0][0];
        }
       

        private List<LeafWithPath> ApplyRowSort(List<LeafWithPath> rowLeaves, PivotByArgs args) =>
    ApplyLeafSort(rowLeaves, args.RowSortOrders);

        private List<LeafWithPath> ApplyColSort(List<LeafWithPath> colLeaves, PivotByArgs args) =>
            ApplyLeafSort(colLeaves, args.ColSortOrders);

        private List<LeafWithPath> ApplyLeafSort(List<LeafWithPath> leaves, int[] sortOrders)
        {
            if (sortOrders == null || sortOrders.All(s => s == 0)) return leaves;

            IOrderedEnumerable<LeafWithPath> ordered = null;
            foreach (var sortOrder in sortOrders)
            {
                if (sortOrder == 0) continue;
                bool desc = sortOrder < 0;
                int col = Math.Abs(sortOrder) - 1;
                var capturedCol = col;

                Func<LeafWithPath, object> keySelector = lp =>
                    capturedCol < lp.Path.Length ? lp.Path[capturedCol] : null;

                if (ordered == null)
                    ordered = desc
                        ? leaves.OrderByDescending(keySelector, _comparer)
                        : leaves.OrderBy(keySelector, _comparer);
                else
                    ordered = desc
                        ? ordered.ThenByDescending(keySelector, _comparer)
                        : ordered.ThenBy(keySelector, _comparer);
            }

            // Bryt oavgjort med efterföljande nyckeldelar
            int maxDepth = leaves.Max(l => l.Path.Length);
            int sortedDepth = sortOrders.Max(s => Math.Abs(s));
            for (int col = sortedDepth; col < maxDepth; col++)
            {
                var capturedCol = col;
                Func<LeafWithPath, object> keySelector = lp =>
                    capturedCol < lp.Path.Length ? lp.Path[capturedCol] : null;
                ordered = ordered.ThenBy(keySelector, _comparer);
            }

            return ordered?.ToList() ?? leaves;
        }

        /// <summary>
        /// Infogar en rad med sammansatt nyckel i ett träd.
        /// Extraherad hjälpmetod – identisk logik som i BuildGroups, delad av GROUPBY och PIVOTBY.
        /// </summary>
        protected GroupLevel InsertIntoTree(
            Dictionary<string, GroupLevel> rootDict,
            List<string> rootOrder,
            object[] keyParts)
        {
            var currentDict = rootDict;
            var currentOrder = rootOrder;
            GroupLevel currentLevel = null;

            for (int depth = 0; depth < keyParts.Length; depth++)
            {
                var keyStr = (keyParts[depth]?.ToString() ?? string.Empty).ToLowerInvariant();
                if (!currentDict.TryGetValue(keyStr, out currentLevel))
                {
                    currentLevel = new GroupLevel { Key = keyParts[depth] };
                    currentDict[keyStr] = currentLevel;
                    currentOrder.Add(keyStr);
                }

                if (depth < keyParts.Length - 1)
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
            return currentLevel; // <-- returnera lövet
        }

        private static string MakePivotKey(object[] parts) =>
            string.Join("\u001F", parts.Select(p => p?.ToString()?.ToLowerInvariant() ?? string.Empty).ToArray());

        private InMemoryRange RenderPivot(
    List<LeafWithPath> rowLeaves,
    List<LeafWithPath> colLeaves,
    Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
    PivotByArgs args,
    ParsingContext context)
        {
            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColKeyRows = args.ColFields.Size.NumberOfCols;
            int nColLeaves = colLeaves.Count;
            int nRowLeaves = rowLeaves.Count;
            int nValCols = args.Values.Size.NumberOfCols;

            bool showRowTotal = args.RowTotalDepth != TotalDepthNoTotals;
            bool showColTotal = args.ColTotalDepth != TotalDepthNoTotals;
            bool rowTotalAtTop = args.RowTotalDepth < 0;
            bool colTotalAtLeft = args.ColTotalDepth < 0;

            int dataRows = nRowLeaves;
            int dataCols = nColLeaves;
            int totalRows = nColKeyRows + dataRows + (showRowTotal ? 1 : 0);
            int totalCols = nRowKeyCols + dataCols + (showColTotal ? 1 : 0);

            var result = new InMemoryRange(totalRows, (short)totalCols);

            int grandTotalCol = colTotalAtLeft ? nRowKeyCols : nRowKeyCols + dataCols;
            int grandTotalRow = rowTotalAtTop ? nColKeyRows : nColKeyRows + dataRows;
            int dataRowStart = nColKeyRows + (rowTotalAtTop ? 1 : 0);
            int colOffset = colTotalAtLeft ? 1 : 0;

            // --- Rubrikrader (en per kolumnnivå) ---
            for (int level = 0; level < nColKeyRows; level++)
            {
                for (int ci = 0; ci < nColLeaves; ci++)
                {
                    var colPath = colLeaves[ci].Path;
                    var val = level < colPath.Length ? colPath[level] : null;
                    result.SetValue(level, nRowKeyCols + colOffset + ci, val);
                }

                if (showColTotal)
                    result.SetValue(level, grandTotalCol, level == 0 ? (object)"Total" : string.Empty);
            }

            // --- Grand total-rad överst ---
            if (rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, nColKeyRows, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, nColLeaves, colOffset, grandTotalCol, showColTotal);

            // --- Datarader ---
            for (int ri = 0; ri < nRowLeaves; ri++)
            {
                int outputRow = dataRowStart + ri;
                var rowPath = rowLeaves[ri].Path;
                var rowLeaf = rowLeaves[ri].Leaf;

                for (int k = 0; k < rowPath.Length; k++)
                    result.SetValue(outputRow, k, rowPath[k]);

                string rowKey = MakePivotKey(rowPath);

                for (int ci = 0; ci < nColLeaves; ci++)
                {
                    string colKey = MakePivotKey(colLeaves[ci].Path);

                    object aggregated = null;
                    if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                        colMap.TryGetValue(colKey, out var cellVals))
                    {
                        aggregated = Aggregate(args.Function, cellVals, context,
                            args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                    }
                    result.SetValue(outputRow, nRowKeyCols + colOffset + ci, aggregated);
                }

                if (showColTotal)
                {
                    var grandTotalVal = rowLeaf.SubtotalValues != null && rowLeaf.SubtotalValues.Count > 0
                        ? rowLeaf.SubtotalValues[0][0]
                        : null;
                    result.SetValue(outputRow, grandTotalCol, grandTotalVal);
                }
            }

            // --- Grand total-rad nederst ---
            if (!rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, grandTotalRow, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, nColLeaves, colOffset, grandTotalCol, showColTotal);

            return result;
        }

        private void WriteGrandTotalRow(
            InMemoryRange result,
            int r,
            List<LeafWithPath> colLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args,
            ParsingContext context,
            int nRowKeyCols,
            int nColLeaves,
            int colOffset,
            int grandTotalCol,
            bool showColTotal)
        {
            result.SetValue(r, 0, "Total");
            for (int c = 1; c < nRowKeyCols; c++)
                result.SetValue(r, c, string.Empty);

            for (int ci = 0; ci < nColLeaves; ci++)
            {
                string colKey = MakePivotKey(colLeaves[ci].Path);

                var allValsForCol = pivotMap.Values
                    .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                        ? cv
                        : Enumerable.Empty<object[]>())
                    .ToList();

                object grandVal = null;
                if (allValsForCol.Count > 0)
                {
                    grandVal = Aggregate(args.Function, allValsForCol, context,
                        args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                }
                result.SetValue(r, nRowKeyCols + colOffset + ci, grandVal);
            }

            // Hörncellen: aggregera hela AllValuesInOrder
            if (showColTotal)
            {
                var colValues = args.AllValuesInOrder
                    .Select(v => new object[] { v[0] })
                    .ToList();
                var cornerVal = Aggregate(args.Function, colValues, context,
                    args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                result.SetValue(r, grandTotalCol, cornerVal);
            }
        }

        private List<LeafWithPath> CollectLeavesWithPath(List<GroupLevel> levels, object[] parentPath)
        {
            var result = new List<LeafWithPath>();
            foreach (var level in levels)
            {
                var path = new object[parentPath.Length + 1];
                for (int i = 0; i < parentPath.Length; i++)
                    path[i] = parentPath[i];
                path[parentPath.Length] = level.Key;

                Debug.WriteLine($"Level key={level.Key}, IsLeaf={level.IsLeaf}, path.Length={path.Length}, path={string.Join("|", path.Select(p => p?.ToString()).ToArray())}");

                if (level.IsLeaf)
                    result.Add(new LeafWithPath(level, path));
                else
                    result.AddRange(CollectLeavesWithPath(level.Children, path));               
            }
            return result;
        }

        private class LeafWithPath
        {
            public GroupLevel Leaf { get; private set; }
            public object[] Path { get; private set; }

            public LeafWithPath(GroupLevel leaf, object[] path)
            {
                Leaf = leaf;
                Path = path;
            }
        }

    }
}