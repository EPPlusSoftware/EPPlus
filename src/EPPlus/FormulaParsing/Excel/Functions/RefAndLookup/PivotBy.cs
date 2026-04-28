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
        public override int ArgumentMinLength => 3; // kan
                                                    // ske ska vara 4
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

        // 6. Räkna t dimensioner: nRows, nCols
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
            else
            {
                args.RowTotalDepth = TotalDepthGrandOnly; // default = 1
            }

            // SortOrder for RowFields (optional)
            if (arguments.Count > 6 && arguments[6].Value != null)
                args.RowSortOrders = ParseSortOrderArg(arguments[6]);

            // TotalDepth for columns (optional, default = 1)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[7], args.ColFields.Size.NumberOfCols, out int colTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.ColTotalDepth = colTotalDepth;
            }
            else
            {
                args.ColTotalDepth = TotalDepthGrandOnly; // default = 1
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
                    currentLevel = new GroupLevel
                    {
                        Key = keyParts[depth],
                        KeyParts = keyParts.Take(depth + 1).ToArray()
                    };
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
            return currentLevel;
        }

        private GroupLevel FindParentInTree(GroupLevel target, List<GroupLevel> levels)
        {
            foreach (var level in levels)
            {
                if (!level.IsLeaf && level.Children.Contains(target))
                    return level;

                if (!level.IsLeaf)
                {
                    var found = FindParentInTree(target, level.Children);
                    if (found != null) return found;
                }
            }
            return null; // toppnivå, ingen förälder
        }

        private List<object[]> ResolveRelativeToValues(
    RelativeTo relativeTo,
    LeafWithPath colLeaf,
    List<LeafWithPath> colLeaves,
    string rowKey,
    Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
        PivotByArgs args)
        {
            switch (relativeTo)
            {
                case RelativeTo.RowTotals: // 1
                    {
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        return colMap.Values
                            .SelectMany(vals => vals)
                            .ToList();
                    }
                case RelativeTo.GrandTotals: // 2
                    {
                        // Nämnare = alla värden i hela datasetet
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

                        // Hämta radens egna värden under alla syskonkolumner
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        return siblingLeaves
                            .Where(leaf => colMap.ContainsKey(MakePivotKey(leaf.Path)))
                            .SelectMany(leaf => colMap[MakePivotKey(leaf.Path)])
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: // 4
                    {
                        // Nämnare = alla raders värden för denna kolumn (kolumnens totalsumma)
                        string colKey = MakePivotKey(colLeaf.Path);
                        return pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv
                                : Enumerable.Empty<object[]>())
                            .ToList();
                    }
                default:
                    return null;
            }
        }
        private List<object[]> ResolveRelativeToValuesForTotal(
    RelativeTo relativeTo,
    LeafWithPath colLeaf,
    List<LeafWithPath> colLeaves,
    Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
        PivotByArgs args)
        {
            switch (relativeTo)
            {
                case RelativeTo.RowTotals: // 1
                    {
                        return pivotMap.Values
                            .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                            .ToList();
                    }
                case RelativeTo.GrandTotals: // 2
                    {
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        // Nämnaren = alla värden under förälderns grupp (alla syskonkolumner)
                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

                        return pivotMap.Values
                            .SelectMany(colMap => siblingLeaves
                                .Where(leaf => colMap.ContainsKey(MakePivotKey(leaf.Path)))
                                .SelectMany(leaf => colMap[MakePivotKey(leaf.Path)]))
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: // 4
                    {
                        // Total-raden: kolumn / kolumn = 1
                        string colKey = MakePivotKey(colLeaf.Path);
                        var colVals = pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv
                                : Enumerable.Empty<object[]>())
                            .ToList();
                        return colVals;
                    }
                default:
                    return null;
            }
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
            int colSubtotalDepth = Math.Abs(args.ColTotalDepth);
            bool showColSubtotals = colSubtotalDepth > 1;

            // Gruppera kolumnlöv per översta nyckelgrupp (nivå 0)
            var colGroups = colLeaves
                .GroupBy(l => l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty)
                .ToList();

            // Bygg ordnad lista av kolumner – löv först, subtotal sist i varje grupp
            var colEntries = new List<ColEntry>();
            foreach (var group in colGroups)
            {
                var groupLeaves = group.ToList();
                foreach (var leaf in groupLeaves)
                    colEntries.Add(new ColEntry { IsSubtotal = false, Leaf = leaf });
                if (showColSubtotals)
                    colEntries.Add(new ColEntry { IsSubtotal = true, GroupKey = group.Key, GroupLeaves = groupLeaves });
            }

            int nDataCols = colEntries.Count;
            int totalRows = nColKeyRows + nRowLeaves + (showRowTotal ? 1 : 0);
            int totalCols = nRowKeyCols + nDataCols + (showColTotal ? 1 : 0);

            var result = new InMemoryRange(totalRows, (short)totalCols);

            int grandTotalCol = colTotalAtLeft ? nRowKeyCols : nRowKeyCols + nDataCols;
            int grandTotalRow = rowTotalAtTop ? nColKeyRows : nColKeyRows + nRowLeaves;
            int dataRowStart = nColKeyRows + (rowTotalAtTop ? 1 : 0);
            int colOffset = colTotalAtLeft ? 1 : 0;

            // --- Rubrikrader ---
            for (int level = 0; level < nColKeyRows; level++)
            {
                int col = nRowKeyCols + colOffset;
                foreach (var entry in colEntries)
                {
                    if (entry.IsSubtotal)
                        result.SetValue(level, col, level == 0 ? entry.GroupLeaves[0].Path[0] : (object)string.Empty);
                    else
                    {
                        var val = level < entry.Leaf.Path.Length ? entry.Leaf.Path[level] : null;
                        result.SetValue(level, col, val);
                    }
                    col++;
                }

                string colTotalLabel = args.ColTotalDepth < 0 ?  "Grand Total" : "Total";
                if (showColTotal)
                    result.SetValue(level, grandTotalCol, level == 0 ? colTotalLabel : string.Empty);
            }

            // --- Grand total-rad överst ---
            if (rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, nColKeyRows, colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, colOffset, grandTotalCol, showColTotal);

            // --- Datarader ---
            for (int ri = 0; ri < nRowLeaves; ri++)
            {
                int outputRow = dataRowStart + ri;
                var rowPath = rowLeaves[ri].Path;
                var rowLeaf = rowLeaves[ri].Leaf;

                for (int k = 0; k < rowPath.Length; k++)
                    result.SetValue(outputRow, k, rowPath[k]);

                string rowKey = MakePivotKey(rowPath);

                int col = nRowKeyCols + colOffset;
                foreach (var entry in colEntries)
                {
                    if (entry.IsSubtotal)
                    {
                        var groupVals = entry.GroupLeaves
                            .SelectMany(l =>
                            {
                                var ck = MakePivotKey(l.Path);
                                if (pivotMap.TryGetValue(rowKey, out var cm) && cm.TryGetValue(ck, out var cv))
                                    return cv;
                                return Enumerable.Empty<object[]>();
                            })
                            .ToList();

                        object subtotalVal = null;
                        if (groupVals.Count > 0)
                            subtotalVal = Aggregate(args.Function, groupVals, context,
                                args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                        result.SetValue(outputRow, col, subtotalVal);
                    }
                    else
                    {
                        string colKey = MakePivotKey(entry.Leaf.Path);
                        object aggregated = null;
                        if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                            colMap.TryGetValue(colKey, out var cellVals))
                        {
                            var relativeToVals = args.Function.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo != RelativeTo.ColumnTotals
                                ? ResolveRelativeToValues(args.RelativeTo, entry.Leaf, colLeaves, rowKey, pivotMap, args)
                                : args.AllValuesInOrder;

                            aggregated = Aggregate(args.Function, cellVals, context, relativeToVals);
                        }
                        result.SetValue(outputRow, col, aggregated);
                    }
                    col++;
                }

                if (showColTotal)
                {
                    var rowAllVals = rowLeaf.Rows.SelectMany(r => r.Values).ToList();
                    var relVals = args.Function.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo != RelativeTo.ColumnTotals
                        ? (args.RelativeTo == RelativeTo.GrandTotals || args.RelativeTo == RelativeTo.ParentRowTotal
                            ? args.AllValuesInOrder
                            : rowAllVals)
                        : args.AllValuesInOrder;
                    var grandTotalVal = Aggregate(args.Function, rowAllVals, context, relVals);
                    result.SetValue(outputRow, grandTotalCol, grandTotalVal);
                }
            }

            // --- Grand total-rad nederst ---
            if (!rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, grandTotalRow, colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, colOffset, grandTotalCol, showColTotal);

            return result;
        }

        private void WriteGrandTotalRow(
    InMemoryRange result,
    int r,
    List<ColEntry> colEntries,
    List<LeafWithPath> colLeaves,
    Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
    PivotByArgs args,
    ParsingContext context,
    int nRowKeyCols,
    int colOffset,
    int grandTotalCol,
    bool showColTotal)
        {
            result.SetValue(r, 0, "Total");
            for (int c = 1; c < nRowKeyCols; c++)
                result.SetValue(r, c, string.Empty);

            int col = nRowKeyCols + colOffset;
            foreach (var entry in colEntries)
            {
                if (entry.IsSubtotal)
                {
                    var groupVals = entry.GroupLeaves
                        .SelectMany(l =>

                        {
                            var ck = MakePivotKey(l.Path);
                            return pivotMap.Values
                                .SelectMany(cm => cm.TryGetValue(ck, out var cv)
                                    ? cv
                                    : Enumerable.Empty<object[]>());
                        })
                        .ToList();

                    object subtotalVal = null;
                    if (groupVals.Count > 0)
                        subtotalVal = Aggregate(args.Function, groupVals, context,
                            args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                    result.SetValue(r, col, subtotalVal);
                }
                else
                {
                    string colKey = MakePivotKey(entry.Leaf.Path);
                    var allValsForCol = pivotMap.Values
                        .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                            ? cv
                            : Enumerable.Empty<object[]>())
                        .ToList();

                    object grandVal = null;
                    if (allValsForCol.Count > 0)
                    {
                        var relativeToVals = args.Function.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo != RelativeTo.ColumnTotals
                            ? ResolveRelativeToValuesForTotal(args.RelativeTo, entry.Leaf, colLeaves, pivotMap, args)
                            : args.AllValuesInOrder;

                        grandVal = Aggregate(args.Function, allValsForCol, context, relativeToVals);
                    }
                    result.SetValue(r, col, grandVal);
                }
                col++;
            }

            if (showColTotal)
            {
                var allVals = args.AllValuesInOrder.Select(v => new object[] { v[0] }).ToList();
                var cornerVal = Aggregate(args.Function, allVals, context,
                    args.Function.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo != RelativeTo.ColumnTotals
                        ? allVals  // dela med sig själv → 1
                        : args.AllValuesInOrder);
                result.SetValue(r, grandTotalCol, cornerVal);
            }

        }

        private class ColEntry
        {
            public bool IsSubtotal { get; set; }
            public string GroupKey { get; set; }
            public List<LeafWithPath> GroupLeaves { get; set; }
            public LeafWithPath Leaf { get; set; }
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
        private List<LeafWithPath> CollectLeaves(List<LeafWithPath> leaves)
        {
            return leaves; // redan platta löv, returnera direkt
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