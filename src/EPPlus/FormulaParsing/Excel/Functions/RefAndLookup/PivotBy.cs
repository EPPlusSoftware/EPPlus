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
    internal partial class PivotBy : GroupByFunctionBase
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
                var v = Convert.ToInt32(arguments[4].Value);
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
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        // En kolumnnivå = ingen riktig föräldragrupp finns.
                        // Faller tillbaka på RowTotals: nämnare = radens totalsumma över ALLA kolumner.
                        if (colLeaf.Path.Length <= 1)
                        {
                            return colMap.Values
                                .SelectMany(vals => vals)
                                .ToList();
                        }

                        // Flernivå: nämnare = radens värden inom förälderns kolumngrupp
                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

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
                        // En kolumnnivå = faller tillbaka på GrandTotals som nämnare
                        if (colLeaf.Path.Length <= 1)
                        {
                            return pivotMap.Values
                                .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                                .ToList();
                        }

                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

                        return pivotMap.Values
                            .SelectMany(cm => siblingLeaves
                                .Where(leaf => cm.ContainsKey(MakePivotKey(leaf.Path)))
                                .SelectMany(leaf => cm[MakePivotKey(leaf.Path)]))
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
            int nRowLeaves = rowLeaves.Count;
            int nFunctions = args.Functions.Count;
            bool isVStack = args.FunctionLayout == FunctionLayout.Vertical;
            bool isHStack = args.FunctionLayout == FunctionLayout.Horizontal;
            bool multipleFunctions = isVStack || isHStack;

            bool showRowTotal = args.RowTotalDepth != TotalDepthNoTotals;
            bool showColTotal = args.ColTotalDepth != TotalDepthNoTotals;
            bool rowTotalAtTop = args.RowTotalDepth < 0;
            bool colTotalAtLeft = args.ColTotalDepth < 0;
            int colSubtotalDepth = Math.Abs(args.ColTotalDepth);
            bool showColSubtotals = colSubtotalDepth > 1;

            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool showFieldHeaders = resolvedHeaders == FieldHeaders.YesAndShow;
            int fieldHeaderRows = showFieldHeaders ? 1 : 0;
            int headerDataRows = showFieldHeaders ? 1 : 0;
            int functionHeaderRows = isHStack ? 1 : 0;
            int functionNameCol = isVStack ? nRowKeyCols : -1;
            int functionColOffset = isVStack ? 1 : 0;

            var colGroups = colLeaves
                .GroupBy(l => l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty)
                .ToList();

            var colEntries = new List<ColEntry>();
            foreach (var group in colGroups)
            {
                var groupLeaves = group.ToList();
                foreach (var leaf in groupLeaves)
                    colEntries.Add(new ColEntry { IsSubtotal = false, Leaf = leaf });
                if (showColSubtotals)
                    colEntries.Add(new ColEntry { IsSubtotal = true, GroupKey = group.Key, GroupLeaves = groupLeaves });
            }

            int colsPerEntry = isHStack ? nFunctions : 1;
            int nDataCols = colEntries.Count * colsPerEntry;
            int nTotalCols = showColTotal ? colsPerEntry : 0;
            int totalCols = nRowKeyCols + functionColOffset + nDataCols + nTotalCols;

            int rowsPerLeaf = isVStack ? nFunctions : 1;
            int rowsPerTotal = isVStack ? nFunctions : 1;
            int totalRows = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                          + nRowLeaves * rowsPerLeaf
                          + (showRowTotal ? rowsPerTotal : 0);

            int dataRowStart = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                             + (rowTotalAtTop ? rowsPerTotal : 0);
            int grandTotalRow = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                              + (rowTotalAtTop ? 0 : nRowLeaves * rowsPerLeaf);

            var result = new InMemoryRange(totalRows, (short)totalCols);

            int dataColStart = nRowKeyCols + functionColOffset;
            int grandTotalCol = colTotalAtLeft ? dataColStart : dataColStart + nDataCols;
            int colOffset = colTotalAtLeft ? colsPerEntry : 0;

            // --- Fältnamnrad ---
            if (showFieldHeaders)
            {
                for (int i = 0; i < args.ColFields.Size.NumberOfCols; i++)
                    result.SetValue(0, dataColStart + colOffset + i, args.ColFields.GetOffset(0, i));
            }

            // --- Rubrikrader ---
            for (int level = 0; level < nColKeyRows; level++)
            {
                int outputLevel = fieldHeaderRows + level;
                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    var val = entry.IsSubtotal
                        ? (level == 0 ? entry.GroupLeaves[0].Path[0] : (object)string.Empty)
                        : (level < entry.Leaf.Path.Length ? entry.Leaf.Path[level] : null);

                    for (int f = 0; f < colsPerEntry; f++)
                    {
                        result.SetValue(outputLevel, col, f == 0 ? val : string.Empty);
                        col++;
                    }
                }

                if (showColTotal)
                {
                    string colTotalLabel = Math.Abs(args.ColTotalDepth) > 1 ? "Grand Total" : "Total";
                    for (int f = 0; f < colsPerEntry; f++)
                    {
                        result.SetValue(outputLevel, col, f == 0 && level == 0 ? colTotalLabel : string.Empty);
                        col++;
                    }
                }
            }

            // --- HSTACK: Funktionsnamnsrad ---
            if (isHStack)
            {
                var functionNames = ResolveFunctionHeaders(args.Functions);
                int functionHeaderRow = fieldHeaderRows + nColKeyRows;
                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    foreach (var name in functionNames)
                    {
                        result.SetValue(functionHeaderRow, col, name);
                        col++;
                    }
                }
                if (showColTotal)
                {
                    foreach (var name in functionNames)
                    {
                        result.SetValue(functionHeaderRow, col, name);
                        col++;
                    }
                }
            }

            // --- Header-datarad ---
            if (showFieldHeaders)
            {
                int headerDataRow = fieldHeaderRows + nColKeyRows + functionHeaderRows;

                for (int i = 0; i < nRowKeyCols; i++)
                    result.SetValue(headerDataRow, i, args.RowFields.GetOffset(0, i));

                var headerValue = args.Values.GetOffset(0, 0);
                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    for (int fc = 0; fc < colsPerEntry; fc++)
                    {
                        if (!entry.IsSubtotal)
                            result.SetValue(headerDataRow, col, headerValue);
                        col++;
                    }
                }

                if (showColTotal)
                {
                    for (int fc = 0; fc < colsPerEntry; fc++)
                    {
                        result.SetValue(headerDataRow, col, headerValue);
                        col++;
                    }
                }
            }

            // --- Grand total-rad överst ---
            if (rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows,
                                   colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, functionNameCol, dataColStart, colOffset,
                                   grandTotalCol, showColTotal, colsPerEntry, isVStack);

            // --- Datarader ---
            for (int ri = 0; ri < nRowLeaves; ri++)
            {
                var rowPath = rowLeaves[ri].Path;
                var rowLeaf = rowLeaves[ri].Leaf;
                string rowKey = MakePivotKey(rowPath);
                var functionNames = ResolveFunctionHeaders(args.Functions);

                for (int fi = 0; fi < nFunctions; fi++)
                {
                    int outputRow = dataRowStart + ri * rowsPerLeaf + fi;
                    var f = args.Functions[fi];

                    for (int k = 0; k < rowPath.Length; k++)
                        result.SetValue(outputRow, k, rowPath[k]);

                    if (isVStack)
                        result.SetValue(outputRow, functionNameCol, functionNames[fi]);

                    int col = dataColStart + colOffset;
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
                                }).ToList();

                            if (isHStack)
                            {
                                foreach (var func in args.Functions)
                                {
                                    object val = groupVals.Count > 0
                                        ? Aggregate(func, groupVals, context,
                                            func.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                        : null;
                                    if (val != null)
                                        result.SetValue(outputRow, col, val);
                                    col++;
                                }
                            }
                            else
                            {
                                object val = groupVals.Count > 0
                                    ? Aggregate(f, groupVals, context,
                                        f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                    : null;
                                if (val != null)
                                    result.SetValue(outputRow, col, val);
                                col++;
                            }
                        }
                        else
                        {
                            string colKey = MakePivotKey(entry.Leaf.Path);

                            if (isHStack)
                            {
                                foreach (var func in args.Functions)
                                {
                                    object aggregated = null;
                                    if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                                        colMap.TryGetValue(colKey, out var cellVals))
                                    {
                                        var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                            ? RelativeTo.ParentRowTotal
                                            : args.RelativeTo;
                                        var relativeToVals = func.EtaFunction?.Name == "PERCENTOF"
                                            ? ResolveRelativeToValues(effectiveRelativeTo, entry.Leaf, colLeaves, rowKey, pivotMap, args) ?? args.AllValuesInOrder
                                            : args.AllValuesInOrder;
                                        aggregated = Aggregate(func, cellVals, context, relativeToVals);
                                    }
                                    if (aggregated != null)
                                        result.SetValue(outputRow, col, aggregated);
                                    col++;
                                }
                            }
                            else
                            {
                                object aggregated = null;
                                if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                                    colMap.TryGetValue(colKey, out var cellVals))
                                {
                                    var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                        ? RelativeTo.ParentRowTotal
                                        : args.RelativeTo;
                                    var relativeToVals = f.EtaFunction?.Name == "PERCENTOF"
                                        ? ResolveRelativeToValues(effectiveRelativeTo, entry.Leaf, colLeaves, rowKey, pivotMap, args) ?? args.AllValuesInOrder
                                        : args.AllValuesInOrder;
                                    aggregated = Aggregate(f, cellVals, context, relativeToVals);
                                }
                                if (aggregated != null)
                                    result.SetValue(outputRow, col, aggregated);
                                col++;
                            }
                        }
                    }

                    if (showColTotal)
                    {
                        var rowAllVals = rowLeaf.Rows.SelectMany(r => r.Values).ToList();
                        if (isHStack)
                        {
                            foreach (var func in args.Functions)
                            {
                                var relVals = args.AllValuesInOrder;
                                if (func.EtaFunction?.Name == "PERCENTOF")
                                {
                                    var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                                        ? RelativeTo.ParentRowTotal
                                        : args.RelativeTo;
                                    relVals = effectiveRelativeTo == RelativeTo.GrandTotals || effectiveRelativeTo == RelativeTo.ParentRowTotal
                                        ? args.AllValuesInOrder
                                        : rowAllVals;
                                }
                                var totalVal = Aggregate(func, rowAllVals, context, relVals);
                                if (totalVal != null)
                                    result.SetValue(outputRow, col, totalVal);
                                col++;
                            }
                        }
                        else
                        {
                            var relVals = args.AllValuesInOrder;
                            if (f.EtaFunction?.Name == "PERCENTOF")
                            {
                                var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                                    ? RelativeTo.ParentRowTotal
                                    : args.RelativeTo;
                                relVals = effectiveRelativeTo == RelativeTo.GrandTotals || effectiveRelativeTo == RelativeTo.ParentRowTotal
                                    ? args.AllValuesInOrder
                                    : rowAllVals;
                            }
                            var totalVal = Aggregate(f, rowAllVals, context, relVals);
                            if (totalVal != null)
                                result.SetValue(outputRow, col, totalVal);
                            col++;
                        }
                    }

                    if (isHStack) break;
                }
            }

            // --- Grand total-rad nederst ---
            if (!rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, grandTotalRow, colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, functionNameCol, dataColStart, colOffset,
                                   grandTotalCol, showColTotal, colsPerEntry, isVStack);

            return result;
        }

        private void WriteGrandTotalRow(
            InMemoryRange result,
            int startRow,
            List<ColEntry> colEntries,
            List<LeafWithPath> colLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args,
            ParsingContext context,
            int nRowKeyCols,
            int functionNameCol,
            int dataColStart,
            int colOffset,
            int grandTotalCol,
            bool showColTotal,
            int colsPerEntry,
            bool isVStack)
        {
            var functionNames = ResolveFunctionHeaders(args.Functions);
            int nFunctions = args.Functions.Count;
            int rowsPerTotal = isVStack ? nFunctions : 1;

            for (int fi = 0; fi < rowsPerTotal; fi++)
            {
                int r = startRow + fi;
                var f = args.Functions[fi];

                result.SetValue(r, 0, "Total");
                for (int c = 1; c < nRowKeyCols; c++)
                    result.SetValue(r, c, string.Empty);

                if (isVStack)
                    result.SetValue(r, functionNameCol, functionNames[fi]);

                int col = dataColStart + colOffset;
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
                                        ? cv : Enumerable.Empty<object[]>());
                            }).ToList();

                        if (isVStack)
                        {
                            object val = groupVals.Count > 0
                                ? Aggregate(f, groupVals, context,
                                    f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                : null;
                            if (val != null)
                                result.SetValue(r, col, val);
                            col++;
                        }
                        else
                        {
                            foreach (var func in args.Functions)
                            {
                                object val = groupVals.Count > 0
                                    ? Aggregate(func, groupVals, context,
                                        func.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                    : null;
                                if (val != null)
                                    result.SetValue(r, col, val);
                                col++;
                            }
                        }
                    }
                    else
                    {
                        string colKey = MakePivotKey(entry.Leaf.Path);
                        var allValsForCol = pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv : Enumerable.Empty<object[]>())
                            .ToList();

                        if (isVStack)
                        {
                            object grandVal = null;
                            if (allValsForCol.Count > 0)
                            {
                                var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                    ? RelativeTo.ParentRowTotal
                                    : args.RelativeTo;
                                var relativeToVals = f.EtaFunction?.Name == "PERCENTOF"
                                    ? ResolveRelativeToValuesForTotal(effectiveRelativeTo, entry.Leaf, colLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                    : args.AllValuesInOrder;
                                grandVal = Aggregate(f, allValsForCol, context, relativeToVals);
                            }
                            if (grandVal != null)
                                result.SetValue(r, col, grandVal);
                            col++;
                        }
                        else
                        {
                            foreach (var func in args.Functions)
                            {
                                object grandVal = null;
                                if (allValsForCol.Count > 0)
                                {
                                    var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                        ? RelativeTo.ParentRowTotal
                                        : args.RelativeTo;
                                    var relativeToVals = func.EtaFunction?.Name == "PERCENTOF"
                                        ? ResolveRelativeToValuesForTotal(effectiveRelativeTo, entry.Leaf, colLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                        : args.AllValuesInOrder;
                                    grandVal = Aggregate(func, allValsForCol, context, relativeToVals);
                                }
                                if (grandVal != null)
                                    result.SetValue(r, col, grandVal);
                                col++;
                            }
                        }
                    }
                }

                if (showColTotal)
                {
                    var allVals = args.AllValuesInOrder.Select(v => new object[] { v[0] }).ToList();
                    if (isVStack)
                    {
                        var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                            ? RelativeTo.ParentRowTotal
                            : args.RelativeTo;
                        var cornerRelVals = f.EtaFunction?.Name == "PERCENTOF"
                            ? (effectiveRelativeTo == RelativeTo.ParentRowTotal || effectiveRelativeTo == RelativeTo.GrandTotals
                                ? allVals : args.AllValuesInOrder)
                            : args.AllValuesInOrder;
                        var cornerVal = Aggregate(f, allVals, context, cornerRelVals);
                        if (cornerVal != null)
                            result.SetValue(r, col, cornerVal);
                        col++;
                    }
                    else
                    {
                        foreach (var func in args.Functions)
                        {
                            var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                ? RelativeTo.ParentRowTotal
                                : args.RelativeTo;
                            var cornerRelVals = func.EtaFunction?.Name == "PERCENTOF"
                                ? (effectiveRelativeTo == RelativeTo.ParentRowTotal || effectiveRelativeTo == RelativeTo.GrandTotals
                                    ? allVals : args.AllValuesInOrder)
                                : args.AllValuesInOrder;
                            var cornerVal = Aggregate(func, allVals, context, cornerRelVals);
                            if (cornerVal != null)
                                result.SetValue(r, col, cornerVal);
                            col++;
                        }
                    }
                }

                if (!isVStack) break;
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

                if (level.IsLeaf)
                    result.Add(new LeafWithPath(level, path));
                else
                    result.AddRange(CollectLeavesWithPath(level.Children, path));               
            }
            return result;
        }        
    }
}